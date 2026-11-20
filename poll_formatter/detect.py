"""Find poll slides in a presentation. Reads only: never modifies a slide."""
from dataclasses import dataclass, field
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.presentation import Presentation
from pptx.shapes.base import BaseShape

from . import spec

PRIMARY_LABEL_PREFIXES = (
    "question",
    "poll",
    "quiz",
    "mcq",
    "quick quiz",
    "multiple choice",
)


@dataclass
class PollSlide:
    number: int                    # 1-based slide number
    label: BaseShape | None        # the "QUESTION"/"QUICK QUIZ"/... box that becomes "POLL N"
    question: BaseShape
    options: list[BaseShape]       # exactly 4, in A, B, C, D order (row first, then column)
    extras: list[BaseShape] = field(default_factory=list)  # other label text, e.g. a 2nd label chip
    existing_letters: list[BaseShape] = field(default_factory=list)  # existing A, B, C, D letter shapes


@dataclass
class Skipped:
    number: int
    reason: str                    # shown to the user, e.g. "label found but 3 options"


@dataclass
class Detection:
    polls: list[PollSlide]
    skipped: list[Skipped]         # near-misses only, not every non-poll slide


def _group_has_text(shape: BaseShape) -> bool:
    """Check if a grouped shape contains any text frames with non-empty text.
    Purely graphical/vector groups (like Canva icons or illustrations) return False.
    """
    if shape.shape_type != MSO_SHAPE_TYPE.GROUP:
        return False
    if not hasattr(shape, "shapes"):
        return False
    for sub in shape.shapes:
        if sub.has_text_frame and sub.text_frame.text.strip():
            return True
        if sub.shape_type == MSO_SHAPE_TYPE.GROUP and _group_has_text(sub):
            return True
    return False


def _is_label(shape: BaseShape) -> bool:
    """Check if a shape represents a quiz/poll label or category chip."""
    if not shape.has_text_frame:
        return False
    t = shape.text_frame.text.strip().lower()
    if not t:
        return False
    if t in spec.LABEL_WORDS:
        return True
    if len(t) <= 35 and not t.endswith(("?", ":")):
        prefixes = (
            "question",
            "poll",
            "quiz",
            "mcq",
            "concept check",
            "check your understanding",
            "practice question",
            "q.",
            "q ",
            "q-",
        )
        return any(t.startswith(p) for p in prefixes)
    return False


def _is_primary_label(shape: BaseShape) -> bool:
    """Check if a label specifically designates the question/poll (vs a category chip)."""
    t = shape.text_frame.text.strip().lower()
    if len(t) <= 35 and not t.endswith(("?", ":")):
        return any(t.startswith(p) for p in PRIMARY_LABEL_PREFIXES)
    return False


def _sort_options_grid(options: list[BaseShape]) -> list[BaseShape]:
    """Sort 4 options in grid order: row first (top to bottom), then column (left to right)."""
    sorted_by_y = sorted(options, key=lambda s: s.top)
    rows: list[list[BaseShape]] = []

    for opt in sorted_by_y:
        placed = False
        for row in rows:
            ref = row[0]
            vert_overlap = min(ref.top + ref.height, opt.top + opt.height) - max(ref.top, opt.top)
            y_diff = abs(opt.top - ref.top)
            if vert_overlap > 0 or y_diff < ref.height * 0.5:
                row.append(opt)
                placed = True
                break
        if not placed:
            rows.append([opt])

    rows.sort(key=lambda r: min(s.top for s in r))
    result: list[BaseShape] = []
    for r in rows:
        r.sort(key=lambda s: s.left)
        result.extend(r)
    return result


def find_polls(prs: Presentation) -> Detection:
    """Find poll slides in prs and near-misses that should be reported as skipped."""
    polls: list[PollSlide] = []
    skipped: list[Skipped] = []

    for idx, slide in enumerate(prs.slides):
        slide_num = idx + 1

        # Check for grouped shapes that genuinely contain text
        has_grouped_text = any(_group_has_text(s) for s in slide.shapes)

        # Collect top-level text shapes
        labels: list[BaseShape] = []
        non_labels: list[BaseShape] = []

        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text.strip()
            if not text:
                continue

            if _is_label(shape):
                labels.append(shape)
            else:
                non_labels.append(shape)

        # Check if 4 shapes are existing letter boxes (A, B, C, D)
        letter_boxes: list[BaseShape] = []
        extras_detected: list[BaseShape] = []

        single_letters = [
            sh for sh in non_labels
            if sh.text_frame.text.strip().upper() in spec.LETTERS
            and len(sh.text_frame.text.strip()) == 1
        ]
        if len(single_letters) == 4:
            letter_boxes = single_letters
            remaining = [sh for sh in non_labels if sh not in letter_boxes]
            if len(remaining) > 5:
                cand_non_labels: list[BaseShape] = []
                for sh in remaining:
                    t = sh.text_frame.text.strip()
                    if t.isdigit() and sh.top > prs.slide_height * 0.8:
                        extras_detected.append(sh)
                    elif any(h in t.lower() for h in ("think it through", "hint")):
                        extras_detected.append(sh)
                    else:
                        cand_non_labels.append(sh)
                remaining = cand_non_labels
            if len(remaining) == 5:
                non_labels = remaining

        # If a candidate slide genuinely has grouped text, report skipped
        if has_grouped_text and (
            labels or any(nl.text_frame.text.strip().endswith(('?', ':')) for nl in non_labels)
        ):
            skipped.append(Skipped(number=slide_num, reason="grouped shapes or SmartArt not supported"))
            continue

        if labels:
            # Pick primary label first
            primary = [l for l in labels if _is_primary_label(l)]
            if primary:
                primary.sort(key=lambda s: s.top)
                main_label = primary[0]
                extras = [l for l in labels if l is not main_label] + extras_detected
            else:
                labels.sort(key=lambda s: s.top)
                main_label = labels[0]
                extras = labels[1:] + extras_detected

            if len(non_labels) == 5:
                non_labels.sort(key=lambda s: s.top)
                question = non_labels[0]
                options = _sort_options_grid(non_labels[1:])
                polls.append(PollSlide(
                    number=slide_num,
                    label=main_label,
                    question=question,
                    options=options,
                    extras=extras,
                    existing_letters=letter_boxes,
                ))
            elif len(non_labels) == 1:
                # Question with 0 options: not a poll, not reported as skipped
                pass
            elif len(non_labels) > 1:
                opt_count = len(non_labels) - 1
                skipped.append(Skipped(
                    number=slide_num,
                    reason=f"label found but {opt_count} options",
                ))
            continue

        # No label words found
        if len(non_labels) == 5:
            non_labels.sort(key=lambda s: s.top)
            candidate_q = non_labels[0]
            q_text = candidate_q.text_frame.text.strip()
            if q_text.endswith(('?', ':')):
                options = _sort_options_grid(non_labels[1:])
                polls.append(PollSlide(
                    number=slide_num,
                    label=None,
                    question=candidate_q,
                    options=options,
                    extras=extras_detected,
                    existing_letters=letter_boxes,
                ))

    return Detection(polls=polls, skipped=skipped)
