"""Restyle poll slides into the format AI Tutor accepts.

Keeps slide background, freeforms, layout, and images.
Only restyles text into target format and updates shape order.
"""
import copy
import re
from pptx.slide import Slide
from pptx.util import Pt

from . import spec
from .detect import PollSlide

NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _clean_option_prefix(tf) -> None:
    """Strip leading option marker like 'A.', 'A)', '(A)', '1.', '1)' from text runs."""
    full_text = tf.text.strip()
    m = re.match(r"^\(?[A-Da-d1-4][\.\)\:\-]\s*", full_text)
    if not m:
        return

    prefix_len = len(m.group(0))
    for p in tf.paragraphs:
        for r in p.runs:
            if prefix_len <= 0:
                break
            if len(r.text) <= prefix_len:
                prefix_len -= len(r.text)
                r.text = ""
            else:
                r.text = r.text[prefix_len:]
                prefix_len = 0


def _ungroup_grp_element(grp_elem) -> list:
    """Ungroup a single <p:grpSp> element into its parent container,
    transforming all child shape coordinates to parent coordinates.
    """
    parent = grp_elem.getparent()
    if parent is None:
        return []

    grpPr = grp_elem.find(f"{{{NS_P}}}grpSpPr")
    if grpPr is None:
        grpPr = grp_elem.find(f"{{{NS_A}}}grpSpPr")

    gx, gy, gcx, gcy = 0, 0, 1, 1
    chx, chy, chcx, chcy = 0, 0, 1, 1
    if grpPr is not None:
        xfrm = grpPr.find(f"{{{NS_A}}}xfrm")
        if xfrm is not None:
            off = xfrm.find(f"{{{NS_A}}}off")
            ext = xfrm.find(f"{{{NS_A}}}ext")
            chOff = xfrm.find(f"{{{NS_A}}}chOff")
            chExt = xfrm.find(f"{{{NS_A}}}chExt")
            if off is not None:
                gx, gy = int(off.get("x", 0)), int(off.get("y", 0))
            if ext is not None:
                gcx, gcy = int(ext.get("cx", 1)), int(ext.get("cy", 1))
            if chOff is not None:
                chx, chy = int(chOff.get("x", 0)), int(chOff.get("y", 0))
            if chExt is not None:
                chcx, chcy = int(chExt.get("cx", 1)) or 1, int(chExt.get("cy", 1)) or 1

    scale_x = gcx / chcx
    scale_y = gcy / chcy

    children = [c for c in list(grp_elem) if not c.tag.endswith(("nvGrpSpPr", "grpSpPr"))]
    idx = parent.index(grp_elem)
    unpacked = []
    for child in children:
        xfrm = None
        for tag in ("spPr", "grpSpPr", "picPr", "cxnSpPr"):
            pr = child.find(f"{{{NS_P}}}{tag}")
            if pr is not None:
                xfrm = pr.find(f"{{{NS_A}}}xfrm")
                if xfrm is not None:
                    break
        if xfrm is None:
            xfrm = child.find(f"{{{NS_P}}}xfrm")
        if xfrm is None:
            xfrm = child.find(f".//{{{NS_A}}}xfrm")

        if xfrm is not None:
            off = xfrm.find(f"{{{NS_A}}}off")
            ext = xfrm.find(f"{{{NS_A}}}ext")
            if off is not None and ext is not None:
                sx = int(off.get("x", 0))
                sy = int(off.get("y", 0))
                scx = int(ext.get("cx", 0))
                scy = int(ext.get("cy", 0))

                abs_x = int(round(gx + (sx - chx) * scale_x))
                abs_y = int(round(gy + (sy - chy) * scale_y))
                abs_cx = int(round(scx * scale_x))
                abs_cy = int(round(scy * scale_y))

                off.set("x", str(abs_x))
                off.set("y", str(abs_y))
                ext.set("cx", str(abs_cx))
                ext.set("cy", str(abs_cy))

        parent.insert(idx, child)
        idx += 1
        unpacked.append(child)

    parent.remove(grp_elem)
    return unpacked


def _ensure_unparented_groups(spTree, shapes: list) -> None:
    """Ensure all shapes in `shapes` have spTree as direct parent by ungrouping their ancestor groups."""
    for sh in shapes:
        if sh is None:
            continue
        elem = sh._element
        while True:
            parent = elem.getparent()
            if parent is None or parent == spTree:
                break
            curr = parent
            while curr.getparent() is not None and curr.getparent() != spTree and curr.getparent().tag.endswith("grpSp"):
                curr = curr.getparent()
            if curr.tag.endswith("grpSp"):
                _ungroup_grp_element(curr)
            else:
                break


def restyle(slide: Slide, poll: PollSlide, poll_no: int, scale: float = 1.0) -> None:
    """Restyle a single poll slide in-place:
    1. Label becomes 'POLL N' in Calibri 18 pt bold (preserving existing color).
    2. Question font becomes Calibri 28.5 pt bold (preserving run colors).
    3. Letters A/B/C/D (Calibri 30 pt bold) are positioned at left of each option.
       Option boxes are nudged right (if not already positioned) so letters do not overlap.
       Option font becomes Calibri 21.75 pt regular (preserving run colors).
    4. Text shapes are moved to the end of the slide's shape tree in reference order:
       label, question, A, optA, B, optB, C, optC, D, optD, extras...
    """
    spTree = slide.shapes._spTree

    # Ensure all poll shapes are direct children of spTree before modifying
    poll_shapes = [poll.label, poll.question] + poll.options + poll.existing_letters + poll.extras
    _ensure_unparented_groups(spTree, poll_shapes)

    # 1. Label
    if poll.label is not None:
        tf = poll.label.text_frame
        while len(tf.paragraphs) > 1:
            p_elem = tf.paragraphs[-1]._p
            p_elem.getparent().remove(p_elem)
        p0 = tf.paragraphs[0]
        if p0.runs:
            r0 = p0.runs[0]
            r0.text = f"POLL {poll_no}"
            r0.font.name = spec.FONT
            r0.font.size = Pt(spec.LABEL[0] * scale)
            r0.font.bold = spec.LABEL[1]
            while len(p0.runs) > 1:
                r_elem = p0.runs[-1]._r
                r_elem.getparent().remove(r_elem)
        else:
            r0 = p0.add_run()
            r0.text = f"POLL {poll_no}"
            r0.font.name = spec.FONT
            r0.font.size = Pt(spec.LABEL[0] * scale)
            r0.font.bold = spec.LABEL[1]
        label_shape = poll.label
    else:
        # Create a label textbox if slide had none
        label_left = poll.question.left
        label_top = max(0, poll.question.top - int(500_000 * scale))
        label_shape = slide.shapes.add_textbox(
            label_left,
            label_top,
            int(3_000_000 * scale),
            int(500_000 * scale),
        )
        p0 = label_shape.text_frame.paragraphs[0]
        r0 = p0.add_run()
        r0.text = f"POLL {poll_no}"
        r0.font.name = spec.FONT
        r0.font.size = Pt(spec.LABEL[0] * scale)
        r0.font.bold = spec.LABEL[1]

    # 2. Question
    for p in poll.question.text_frame.paragraphs:
        for r in p.runs:
            r.font.name = spec.FONT
            r.font.size = Pt(spec.QUESTION[0] * scale)
            r.font.bold = spec.QUESTION[1]
        if not p.runs and p.text:
            p.font.name = spec.FONT
            p.font.size = Pt(spec.QUESTION[0] * scale)
            p.font.bold = spec.QUESTION[1]

    # 3. Letters & Options
    letter_shapes = []
    letter_box_size = int(spec.LETTER_BOX * scale)
    shift = int(1_100_000 * scale)

    has_existing_letters = len(poll.existing_letters) == 4

    for idx, (opt, letter) in enumerate(zip(poll.options, spec.LETTERS)):
        # Clean any leading 'A.', '(A)', etc. from option text
        _clean_option_prefix(opt.text_frame)

        # Extract color element from option runs if present
        opt_color_elem = None
        for p in opt.text_frame.paragraphs:
            for r in p.runs:
                fill = r._r.get_or_add_rPr().find(
                    "{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill"
                )
                if fill is not None:
                    opt_color_elem = fill
                    break
            if opt_color_elem is not None:
                break

        if has_existing_letters:
            let_shape = poll.existing_letters[idx]
            let_tf = let_shape.text_frame
            let_tf.text = letter
            let_r = let_tf.paragraphs[0].runs[0]
            let_r.font.name = spec.FONT
            let_r.font.size = Pt(spec.LETTER[0] * scale)
            let_r.font.bold = spec.LETTER[1]
            if opt_color_elem is not None:
                let_r._r.get_or_add_rPr().append(copy.deepcopy(opt_color_elem))
        else:
            # Create letter text box at original opt position
            let_shape = slide.shapes.add_textbox(
                opt.left,
                opt.top,
                letter_box_size,
                letter_box_size,
            )
            let_shape.text_frame.margin_left = 0
            let_shape.text_frame.margin_top = 0
            let_shape.text_frame.margin_right = 0
            let_shape.text_frame.margin_bottom = 0

            let_p = let_shape.text_frame.paragraphs[0]
            let_r = let_p.add_run()
            let_r.text = letter
            let_r.font.name = spec.FONT
            let_r.font.size = Pt(spec.LETTER[0] * scale)
            let_r.font.bold = spec.LETTER[1]
            if opt_color_elem is not None:
                let_r._r.get_or_add_rPr().append(copy.deepcopy(opt_color_elem))

            # Nudge option box to the right
            opt.left = opt.left + shift
            opt.width = max(opt.width - shift, int(1_000_000 * scale))

        # Restyle option text
        for p in opt.text_frame.paragraphs:
            for r in p.runs:
                r.font.name = spec.FONT
                r.font.size = Pt(spec.OPTION[0] * scale)
                r.font.bold = spec.OPTION[1]
            if not p.runs and p.text:
                p.font.name = spec.FONT
                p.font.size = Pt(spec.OPTION[0] * scale)
                p.font.bold = spec.OPTION[1]

        letter_shapes.append(let_shape)

    # 4. Shape order: label -> question -> A -> optA -> ... -> D -> optD -> extras
    ordered_shapes = [label_shape, poll.question]
    for let, opt in zip(letter_shapes, poll.options):
        ordered_shapes.append(let)
        ordered_shapes.append(opt)
    for ext in poll.extras:
        ordered_shapes.append(ext)

    for sh in ordered_shapes:
        spTree.append(sh._element)

