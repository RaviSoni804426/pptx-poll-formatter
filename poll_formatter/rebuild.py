"""Rebuild extracted polls into the proven AI Tutor presentation format.

Uses the standardized 23-shape blueprint accepted by AI Tutor:
- Shape 0: POLL N label at (960144, 685817)
- Shape 1-3: Question container card, circle accent, bulb/star icon
- Shape 4: Question text box
- Shapes 5-8: Option A container, circle, letter A, option text
- Shapes 9-12: Option B container, circle, letter B, option text
- Shapes 13-16: Option C container, circle, letter C, option text
- Shapes 17-20: Option D container, circle, letter D, option text
- Shape 21: Hint ('Think it through — which one is correct?')
- Shape 22: Slide / page number N
"""
import io
from pathlib import Path
import re
import zipfile
from lxml import etree

from .detect import PollSlide

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"

_TEMPLATE_PATH = Path(__file__).parent / "template.pptx"


def clean_option_text(text: str) -> str:
    """Clean leading option markers like 'A.', '(A)', 'A)', '1.', '1)'."""
    cleaned = re.sub(r"^\(?[A-Da-d1-4][\.\)\:\-]\s*", "", text.strip())
    return cleaned if cleaned else text.strip()


def build_polls_deck(polls: list[PollSlide], template_bytes: bytes | None = None) -> bytes:
    """Build a standardized AI Tutor deck containing only the extracted polls."""
    if template_bytes is None:
        template_bytes = _TEMPLATE_PATH.read_bytes()

    with zipfile.ZipFile(io.BytesIO(template_bytes)) as z_in:
        template_files = {name: z_in.read(name) for name in z_in.namelist()}
        base_slide_xml = template_files["ppt/slides/slide1.xml"]
        base_slide_rels = template_files["ppt/slides/_rels/slide1.xml.rels"]

    out_buf = io.BytesIO()
    with zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED) as z_out:
        sld_id_elems = []
        pres_rel_elems = []
        ct_overrides = []

        for i, poll in enumerate(polls, 1):
            s_tree = etree.fromstring(base_slide_xml)
            spTree = s_tree.find(f".//{{{NS_P}}}spTree")
            shapes = spTree.findall(f"{{{NS_P}}}sp")

            def set_text(sp, val):
                t_elems = sp.findall(f".//{{{NS_A}}}t")
                for t in t_elems[1:]:
                    t.text = ""
                if t_elems:
                    t_elems[0].text = val

            text_shapes = [sp for sp in shapes if sp.findall(f".//{{{NS_A}}}t")]
            # Shape index mapping in the standard 23-shape blueprint:
            # text_shapes[0]: POLL N
            # text_shapes[1]: Question
            # text_shapes[2]: 'A'
            # text_shapes[3]: Option A
            # text_shapes[4]: 'B'
            # text_shapes[5]: Option B
            # text_shapes[6]: 'C'
            # text_shapes[7]: Option C
            # text_shapes[8]: 'D'
            # text_shapes[9]: Option D
            # text_shapes[10]: Hint
            # text_shapes[11]: Slide number
            set_text(text_shapes[0], f"POLL {i}")
            set_text(text_shapes[1], poll.question.text_frame.text.strip())
            set_text(text_shapes[3], clean_option_text(poll.options[0].text_frame.text))
            set_text(text_shapes[5], clean_option_text(poll.options[1].text_frame.text))
            set_text(text_shapes[7], clean_option_text(poll.options[2].text_frame.text))
            set_text(text_shapes[9], clean_option_text(poll.options[3].text_frame.text))
            set_text(text_shapes[11], str(i))

            z_out.writestr(f"ppt/slides/slide{i}.xml", etree.tostring(s_tree))
            z_out.writestr(f"ppt/slides/_rels/slide{i}.xml.rels", base_slide_rels)

            r_id = f"rId{100+i}"
            sld_id_elems.append(
                etree.Element(f"{{{NS_P}}}sldId", attrib={"id": str(256 + i), f"{{{NS_R}}}id": r_id})
            )
            pres_rel_elems.append(
                etree.Element(
                    "Relationship",
                    attrib={
                        "Id": r_id,
                        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                        "Target": f"slides/slide{i}.xml",
                    },
                )
            )
            ct_overrides.append(
                etree.Element(
                    "Override",
                    attrib={
                        "PartName": f"/ppt/slides/slide{i}.xml",
                        "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
                    },
                )
            )

        # Update presentation.xml
        pres_tree = etree.fromstring(template_files["ppt/presentation.xml"])
        sldIdLst = pres_tree.find(f"{{{NS_P}}}sldIdLst")
        sldIdLst.clear()
        for el in sld_id_elems:
            sldIdLst.append(el)
        z_out.writestr("ppt/presentation.xml", etree.tostring(pres_tree))

        # Update presentation.xml.rels
        pres_rels_tree = etree.fromstring(template_files["ppt/_rels/presentation.xml.rels"])
        for rel in list(pres_rels_tree):
            if rel.get("Type").endswith("/slide"):
                pres_rels_tree.remove(rel)
        for el in pres_rel_elems:
            pres_rels_tree.append(el)
        z_out.writestr("ppt/_rels/presentation.xml.rels", etree.tostring(pres_rels_tree))

        # Update [Content_Types].xml
        ct_tree = etree.fromstring(template_files["[Content_Types].xml"])
        for ov in list(ct_tree):
            if ov.get("PartName", "").startswith("/ppt/slides/slide"):
                ct_tree.remove(ov)
        for el in ct_overrides:
            ct_tree.append(el)
        z_out.writestr("[Content_Types].xml", etree.tostring(ct_tree))

        # Copy non-slide files
        for name, content in template_files.items():
            if not name.startswith("ppt/slides/") and name not in [
                "ppt/presentation.xml",
                "ppt/_rels/presentation.xml.rels",
                "[Content_Types].xml",
            ]:
                z_out.writestr(name, content)

    return out_buf.getvalue()


_ICON_SRC = "ppt/media/poll-icon-star.png"
_ICON_DST = "ppt/media/tuto-poll-icon-star.png"


def _fill_blueprint_slide(base_slide_xml: bytes, poll: PollSlide, poll_no: int) -> bytes:
    """Return the proven blueprint slide XML with this poll's content filled in."""
    tree = etree.fromstring(base_slide_xml)
    spTree = tree.find(f".//{{{NS_P}}}spTree")
    text_sps = [sp for sp in spTree.findall(f"{{{NS_P}}}sp")
                if sp.findall(f".//{{{NS_A}}}t")]

    def set_text(sp, val):
        ts = sp.findall(f".//{{{NS_A}}}t")
        for t in ts[1:]:
            t.text = ""
        if ts:
            ts[0].text = val

    set_text(text_sps[0], f"POLL {poll_no}")
    set_text(text_sps[1], poll.question.text_frame.text.strip())
    for k in range(4):
        set_text(text_sps[3 + k * 2],
                 clean_option_text(poll.options[k].text_frame.text))
    set_text(text_sps[11], str(poll_no))
    return etree.tostring(tree)


def build_inplace_deck(pptx_bytes: bytes, polls: list[PollSlide],
                       template_bytes: bytes | None = None) -> bytes:
    """Return the original deck with every poll slide REPLACED by the proven
    AI Tutor blueprint slide.

    Non-poll slides, layouts, masters and media stay byte-identical. This trades
    the source deck's custom poll artwork for a layout AI Tutor is known to
    import, which in-place restyling could not achieve.
    """
    if template_bytes is None:
        template_bytes = _TEMPLATE_PATH.read_bytes()

    with zipfile.ZipFile(io.BytesIO(template_bytes)) as z_t:
        base_slide_xml = z_t.read("ppt/slides/slide1.xml")
        icon_bytes = z_t.read(_ICON_SRC)

    by_number = {p.number: (i + 1, p) for i, p in enumerate(polls)}

    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(pptx_bytes)) as z_in:
        # each poll slide keeps pointing at the layout it already used
        layouts: dict[int, str] = {}
        for n in by_number:
            rels = z_in.read(f"ppt/slides/_rels/slide{n}.xml.rels").decode()
            m = re.search(r'Target="(\.\./slideLayouts/[^"]+)"', rels)
            layouts[n] = m.group(1) if m else "../slideLayouts/slideLayout1.xml"

        with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as z_out:
            for item in z_in.infolist():
                name = item.filename
                m_slide = re.fullmatch(r"ppt/slides/slide(\d+)\.xml", name)
                m_rels = re.fullmatch(r"ppt/slides/_rels/slide(\d+)\.xml\.rels", name)

                if m_slide and int(m_slide.group(1)) in by_number:
                    poll_no, poll = by_number[int(m_slide.group(1))]
                    z_out.writestr(item, _fill_blueprint_slide(base_slide_xml, poll, poll_no))
                elif m_rels and int(m_rels.group(1)) in by_number:
                    n = int(m_rels.group(1))
                    z_out.writestr(item, (
                        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                        f'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/'
                        f'officeDocument/2006/relationships/image" Target="../media/'
                        f'{_ICON_DST.split("/")[-1]}"/>'
                        f'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/'
                        f'officeDocument/2006/relationships/slideLayout" Target="{layouts[n]}"/>'
                        '</Relationships>'
                    ).encode())
                else:
                    z_out.writestr(item, z_in.read(name))

            z_out.writestr(_ICON_DST, icon_bytes)

    return out.getvalue()
