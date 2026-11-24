"""Find poll slides in a presentation. Reads only: never modifies a slide."""
from dataclasses import dataclass, field
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.presentation import Presentation
from pptx.shapes.base import BaseShape

from . import spec

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"

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


def get_absolute_bounds(shape: BaseShape) -> tuple[int, int, int, int]:
    """Get absolute (left, top, width, height) in slide coordinates (EMU),
    properly transforming through any ancestor group shapes.
    """
    elem = shape._element
    xfrm = elem.find(f".//{{{NS_A}}}xfrm")
    if xfrm is None:
        return shape.left, shape.top, shape.width, shape.height
    off = xfrm.find(f"{{{NS_A}}}off")
    ext = xfrm.find(f"{{{NS_A}}}ext")
    x = int(off.get("x", 0)) if off is not None else shape.left
    y = int(off.get("y", 0)) if off is not None else shape.top
    cx = int(ext.get("cx", 0)) if ext is not None else shape.width
    cy = int(ext.get("cy", 0)) if ext is not None else shape.height

    curr = elem.getparent()
    while curr is not None and not curr.tag.endswith("spTree"):
        if curr.tag.endswith("grpSp"):
            grpPr = curr.find(f"{{{NS_P}}}grpSpPr")
            if grpPr is None:
                grpPr = curr.find(f"{{{NS_A}}}grpSpPr")
            if grpPr is not None:
                g_xfrm = grpPr.find(f"{{{NS_A}}}xfrm")
                if g_xfrm is not None:
                    g_off = g_xfrm.find(f"{{{NS_A}}}off")
                    g_ext = g_xfrm.find(f"{{{NS_A}}}ext")
                    g_chOff = g_xfrm.find(f"{{{NS_A}}}chOff")
                    g_chExt = g_xfrm.find(f"{{{NS_A}}}chExt")
                    if g_off is not None and g_ext is not None and g_chOff is not None and g_chExt is not None:
                        gx = int(g_off.get("x", 0))
                        gy = int(g_off.get("y", 0))
                        gcx = int(g_ext.get("cx", 1))
                        gcy = int(g_ext.get("cy", 1))
                        chx = int(g_chOff.get("x", 0))
                        chy = int(g_chOff.get("y", 0))
                        chcx = int(g_chExt.get("cx", 1)) or 1
                        chcy = int(g_chExt.get("cy", 1)) or 1

                        scale_x = gcx / chcx
                        scale_y = gcy / chcy
                        x = gx + (x - chx) * scale_x
                        y = gy + (y - chy) * scale_y
                        cx = cx * scale_x
                        cy = cy * scale_y
        curr = curr.getparent()
    return int(round(x)), int(round(y)), int(round(cx)), int(round(cy))


def _collect_text_shapes(shapes_collection) -> list[BaseShape]:
    """Recursively collect all shapes with non-empty text frames."""
    res: list[BaseShape] = []
    for s in shapes_collection:
        if s.has_text_frame and s.text_frame.text.strip():
            res.append(s)
        if s.shape_type == MSO_SHAPE_TYPE.GROUP and hasattr(s, "shapes"):
            res.extend(_collect_text_shapes(s.shapes))
    return res


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
    sorted_by_y = sorted(options, key=lambda s: get_absolute_bounds(s)[1])
    rows: list[list[BaseShape]] = []

    for opt in sorted_by_y:
        opt_x, opt_y, opt_w, opt_h = get_absolute_bounds(opt)
        placed = False
        for row in rows:
            ref = row[0]
            ref_x, ref_y, ref_w, ref_h = get_absolute_bounds(ref)
            vert_overlap = min(ref_y + ref_h, opt_y + opt_h) - max(ref_y, opt_y)
            y_diff = abs(opt_y - ref_y)
            if vert_overlap > 0 or y_diff < ref_h * 0.5:
                row.append(opt)
                placed = True
                break
        if not placed:
            rows.append([opt])

    rows.sort(key=lambda r: min(get_absolute_bounds(s)[1] for s in r))
    result: list[BaseShape] = []
    for r in rows:
        r.sort(key=lambda s: get_absolute_bounds(s)[0])
        result.extend(r)
    return result


def _detect_candidate_poll(
    shapes: list[BaseShape],
    slide_num: int,
    slide_height: int,
) -> tuple[PollSlide | None, Skipped | None]:
    """Attempt to detect a poll slide from a set of text shapes."""
    labels: list[BaseShape] = []
    non_labels: list[BaseShape] = []

    for shape in shapes:
        if not shape.has_text_frame:
            continue
        text = shape.text_frame.text.strip()
        if not text:
            continue

        if _is_label(shape):
            labels.append(shape)
        else:
            non_labels.append(shape)

    letter_boxes: list[BaseShape] = []
    extras_detected: list[BaseShape] = []

    # Check if 4 shapes are existing letter boxes (A, B, C, D)
    single_letters = [
        sh for sh in non_labels
        if sh.text_frame.text.strip().upper() in spec.LETTERS
        and len(sh.text_frame.text.strip()) == 1
    ]
    if len(single_letters) == 4:
        let_map = {sh.text_frame.text.strip().upper(): sh for sh in single_letters}
        if len(let_map) == 4:
            letter_boxes = [let_map[l] for l in spec.LETTERS]
        else:
            letter_boxes = _sort_options_grid(single_letters)
        remaining = [sh for sh in non_labels if sh not in single_letters]
    else:
        remaining = list(non_labels)

    # Filter out extraneous text like slide numbers and hints if too many candidates
    if len(remaining) > 5:
        cand_non_labels: list[BaseShape] = []
        for sh in remaining:
            t = sh.text_frame.text.strip()
            y = get_absolute_bounds(sh)[1]
            if t.isdigit() and (y > slide_height * 0.75 or y < slide_height * 0.15):
                extras_detected.append(sh)
            elif any(h in t.lower() for h in ("think it through", "hint")):
                extras_detected.append(sh)
            else:
                cand_non_labels.append(sh)
        remaining = cand_non_labels

    if len(remaining) == 5:
        non_labels = remaining

    if labels:
        primary = [l for l in labels if _is_primary_label(l)]
        if primary:
            primary.sort(key=lambda s: get_absolute_bounds(s)[1])
            main_label = primary[0]
            extras = [l for l in labels if l is not main_label] + extras_detected
        else:
            labels.sort(key=lambda s: get_absolute_bounds(s)[1])
            main_label = labels[0]
            extras = labels[1:] + extras_detected

        if len(non_labels) == 5:
            non_labels.sort(key=lambda s: get_absolute_bounds(s)[1])
            question = non_labels[0]
            options = _sort_options_grid(non_labels[1:])
            return PollSlide(
                number=slide_num,
                label=main_label,
                question=question,
                options=options,
                extras=extras,
                existing_letters=letter_boxes,
            ), None
        elif len(non_labels) == 1:
            return None, None
        elif len(non_labels) > 1:
            opt_count = len(non_labels) - 1
            return None, Skipped(
                number=slide_num,
                reason=f"label found but {opt_count} options",
            )
        return None, None

    # No label words found
    if len(non_labels) == 5:
        non_labels.sort(key=lambda s: get_absolute_bounds(s)[1])
        candidate_q = non_labels[0]
        q_text = candidate_q.text_frame.text.strip()
        if q_text.endswith(("?", ":")):
            options = _sort_options_grid(non_labels[1:])
            return PollSlide(
                number=slide_num,
                label=None,
                question=candidate_q,
                options=options,
                extras=extras_detected,
                existing_letters=letter_boxes,
            ), None

    return None, None


def find_polls(prs: Presentation) -> Detection:
    """Find poll slides in prs and near-misses that should be reported as skipped."""
    polls: list[PollSlide] = []
    skipped: list[Skipped] = []

    for idx, slide in enumerate(prs.slides):
        slide_num = idx + 1

        top_shapes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
        poll, skip = _detect_candidate_poll(top_shapes, slide_num, prs.slide_height)
        if poll is not None:
            polls.append(poll)
            continue

        # If top-level shapes did not match a poll, check all shapes including grouped ones
        all_shapes = _collect_text_shapes(slide.shapes)
        if len(all_shapes) != len(top_shapes):
            poll_grouped, skip_grouped = _detect_candidate_poll(all_shapes, slide_num, prs.slide_height)
            if poll_grouped is not None:
                polls.append(poll_grouped)
                continue
            if skip_grouped is not None:
                skipped.append(skip_grouped)
                continue

        if skip is not None:
            skipped.append(skip)

    return Detection(polls=polls, skipped=skipped)

