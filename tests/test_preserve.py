import io
import os
from pathlib import Path
import zipfile
import pytest
import pptx
from lxml import etree

from poll_formatter.detect import find_polls
from poll_formatter.pipeline import process

_LOCAL_TEST_DECK = Path("test_decks") / "Into the Night Sky!_06_AS_CBSE revamped.pptx"
_DOWNLOADS_DECK = Path.home() / "Downloads" / "Into the Night Sky!_06_AS_CBSE revamped.pptx"
_DEFAULT_DECK = _LOCAL_TEST_DECK if _LOCAL_TEST_DECK.exists() else _DOWNLOADS_DECK

NIGHT_SKY = Path(os.environ.get("NIGHT_SKY_DECK", _DEFAULT_DECK))
needs_deck = pytest.mark.skipif(not NIGHT_SKY.exists(), reason=f"missing {NIGHT_SKY}")


@needs_deck
def test_non_poll_slides_byte_identical_and_poll_non_text_shapes_identical():
    with open(NIGHT_SKY, "rb") as f:
        in_bytes = f.read()

    prs_in = pptx.Presentation(io.BytesIO(in_bytes))
    detection = find_polls(prs_in)
    poll_slide_nums = {p.number for p in detection.polls}
    assert poll_slide_nums == {8, 9, 12, 19, 20, 21, 24, 25, 26, 27}

    # Extract original non-text shapes on all poll slides
    orig_non_text_shapes = {}
    for p in detection.polls:
        slide = prs_in.slides[p.number - 1]
        orig_non_text_shapes[p.number] = [
            etree.tostring(sh._element)
            for sh in slide.shapes
            if not sh.has_text_frame or not sh.text_frame.text.strip()
        ]

    # Run full pipeline
    out_bytes, report = process(in_bytes)
    assert report.total_slides == 30
    assert report.poll_slides == [8, 9, 12, 19, 20, 21, 24, 25, 26, 27]

    # Verify non-poll slide XML parts are byte-identical
    poll_part_names = {
        f"ppt/slides/slide{p_num}.xml" for p_num in poll_slide_nums
    }

    with zipfile.ZipFile(io.BytesIO(in_bytes)) as z_in, zipfile.ZipFile(io.BytesIO(out_bytes)) as z_out:
        in_names = set(z_in.namelist())
        out_names = set(z_out.namelist())
        assert in_names == out_names

        for name in in_names:
            if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
                if name not in poll_part_names:
                    # Must be 100% byte-identical
                    assert z_in.read(name) == z_out.read(name), f"Non-poll slide {name} differed!"
            elif not name.startswith("ppt/slides/slide"):
                # All layouts, master slides, media, theme, rels must also be byte-identical
                assert z_in.read(name) == z_out.read(name), f"Non-slide resource {name} differed!"

    # Verify non-text shapes on poll slides in output are identical
    prs_out = pptx.Presentation(io.BytesIO(out_bytes))
    for p_num in poll_slide_nums:
        slide_out = prs_out.slides[p_num - 1]
        out_non_text = [
            etree.tostring(sh._element)
            for sh in slide_out.shapes
            if not sh.has_text_frame or not sh.text_frame.text.strip()
        ]
        assert orig_non_text_shapes[p_num] == out_non_text, f"Non-text shapes on slide {p_num} differed!"
