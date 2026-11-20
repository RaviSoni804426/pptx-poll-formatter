import io
from pathlib import Path
import pytest
import pptx

from poll_formatter.pipeline import process

_LOCAL_TEST_DECK = Path("test_decks") / "Into the Night Sky!_06_AS_CBSE revamped.pptx"


def test_process_corrupt_file_raises_value_error():
    with pytest.raises(ValueError, match="Could not open PPTX file"):
        process(b"not a valid pptx file content")


def test_process_presentation_without_polls():
    prs = pptx.Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    tb = s.shapes.add_textbox(0, 0, 1000, 1000)
    tb.text_frame.text = "Just a title"

    buf = io.BytesIO()
    prs.save(buf)
    in_bytes = buf.getvalue()

    out_bytes, report = process(in_bytes)
    assert report.total_slides == 1
    assert report.poll_slides == []
    assert report.poll_count == 0
    assert report.non_poll_count == 1
    assert report.skipped == []
    assert out_bytes == in_bytes


def test_process_near_miss_reported_in_pipeline():
    prs = pptx.Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    tb_l = s.shapes.add_textbox(0, 0, 1000, 1000)
    tb_l.text_frame.text = "QUIZ"
    tb_q = s.shapes.add_textbox(0, 1000, 1000, 1000)
    tb_q.text_frame.text = "Which planet is red?"
    for i in range(3):
        tb_o = s.shapes.add_textbox(0, 2000 + i * 1000, 1000, 1000)
        tb_o.text_frame.text = f"Option {i+1}"

    buf = io.BytesIO()
    prs.save(buf)

    out_bytes, report = process(buf.getvalue())
    assert report.total_slides == 1
    assert report.poll_slides == []
    assert len(report.skipped) == 1
    assert report.skipped[0].number == 1
    assert "3 options" in report.skipped[0].reason
