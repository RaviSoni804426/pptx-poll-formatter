import os
from pathlib import Path
import pytest
import pptx
from pptx.util import Emu

from poll_formatter.detect import find_polls
from poll_formatter.pipeline import process

import io
import os
from pathlib import Path
import pytest
import pptx
from pptx.util import Inches, Emu, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

from poll_formatter import spec
from poll_formatter.detect import find_polls
from poll_formatter.pipeline import process

DISCOUNT_DECK = Path("test_decks") / "Discount in Shopping.pptx"
needs_discount = pytest.mark.skipif(not DISCOUNT_DECK.exists(), reason=f"missing {DISCOUNT_DECK}")


def test_synthetic_canva_grouped_option_cards_detected_and_restyled():
    prs = pptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Top-level question
    q = slide.shapes.add_textbox(Inches(1), Inches(0.8), Inches(8), Inches(1))
    q.text = "What is the principal amount?"

    # 4 grouped option cards (Canva pattern: card background shape + text box grouped)
    coords = [
        (Inches(1), Inches(2.2), "A. $1,000"),
        (Inches(5), Inches(2.2), "B. $2,000"),
        (Inches(1), Inches(4.0), "C. $3,000"),
        (Inches(5), Inches(4.0), "D. $4,000"),
    ]
    for x, y, text in coords:
        grp = slide.shapes.add_group_shape()
        grp.shapes.add_shape(MSO_SHAPE_TYPE.AUTO_SHAPE, x, y, Inches(3.5), Inches(1.2))
        tb = grp.shapes.add_textbox(x + Inches(0.2), y + Inches(0.2), Inches(3.1), Inches(0.8))
        tb.text = text

    buf = io.BytesIO()
    prs.save(buf)
    in_bytes = buf.getvalue()

    # Process pipeline
    out_bytes, report = process(in_bytes)
    assert report.total_slides == 1
    assert report.poll_slides == [1]
    assert report.poll_count == 1
    assert report.skipped == []

    # Check restyled output presentation
    out_prs = pptx.Presentation(io.BytesIO(out_bytes))
    out_slide = out_prs.slides[0]

    all_texts = [s.text_frame.text.strip() for s in out_slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
    assert all_texts == [
        "POLL 1",
        "What is the principal amount?",
        "A",
        "$1,000",
        "B",
        "$2,000",
        "C",
        "$3,000",
        "D",
        "$4,000",
        "Think it through \u2014 which one is correct?",
        "1",
    ]


def test_synthetic_grouped_question_and_options():
    prs = pptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Question grouped with an icon
    grp_q = slide.shapes.add_group_shape()
    grp_q.shapes.add_shape(MSO_SHAPE_TYPE.AUTO_SHAPE, Inches(0.5), Inches(0.5), Inches(0.5), Inches(0.5))
    tb_q = grp_q.shapes.add_textbox(Inches(1.2), Inches(0.5), Inches(8), Inches(1))
    tb_q.text = "Calculate compound interest for 2 years:"

    # 4 options in a single container group
    grp_opts = slide.shapes.add_group_shape()
    for i, t in enumerate(["₹100", "₹210", "₹250", "₹300"]):
        tb = grp_opts.shapes.add_textbox(Inches(1), Inches(2.0 + i * 1.0), Inches(5), Inches(0.8))
        tb.text = t

    buf = io.BytesIO()
    prs.save(buf)

    out_bytes, report = process(buf.getvalue())
    assert report.total_slides == 1
    assert report.poll_slides == [1]
    assert report.skipped == []


def test_synthetic_slide_with_grouped_diagram_not_falsely_skipped():
    """Non-poll explanation slide with a grouped diagram/equation is treated as normal slide."""
    prs = pptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    tb_title = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(8), Inches(1))
    tb_title.text = "Compound Interest Formula Derivation"

    # Grouped equation diagram
    grp = slide.shapes.add_group_shape()
    t1 = grp.shapes.add_textbox(Inches(1), Inches(2), Inches(4), Inches(1))
    t1.text = "A = P(1 + r/n)^(nt)"
    t2 = grp.shapes.add_textbox(Inches(1), Inches(3.5), Inches(4), Inches(1))
    t2.text = "Where P is principal"

    buf = io.BytesIO()
    prs.save(buf)

    out_bytes, report = process(buf.getvalue())
    assert report.total_slides == 1
    assert report.poll_slides == []
    # Must NOT report "grouped shapes or SmartArt not supported"
    assert report.skipped == []
    assert report.non_poll_count == 1


def test_synthetic_grouped_near_miss_reports_accurate_option_count():
    """Near-miss with grouped shapes reports option count, not grouped shapes error."""
    prs = pptx.Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Quiz label
    tb_lbl = slide.shapes.add_textbox(Inches(1), Inches(0.5), Inches(3), Inches(0.5))
    tb_lbl.text = "QUESTION"

    tb_q = slide.shapes.add_textbox(Inches(1), Inches(1.2), Inches(8), Inches(1))
    tb_q.text = "What is the rate?"

    # Only 3 options inside groups
    for i in range(3):
        grp = slide.shapes.add_group_shape()
        grp.shapes.add_shape(MSO_SHAPE_TYPE.AUTO_SHAPE, Inches(1), Inches(2.5 + i * 1.2), Inches(3), Inches(1))
        tb = grp.shapes.add_textbox(Inches(1), Inches(2.5 + i * 1.2), Inches(3), Inches(1))
        tb.text = f"Option {i+1}"

    buf = io.BytesIO()
    prs.save(buf)

    out_bytes, report = process(buf.getvalue())
    assert report.total_slides == 1
    assert report.poll_slides == []
    assert len(report.skipped) == 1
    assert report.skipped[0].number == 1
    assert "3 options" in report.skipped[0].reason
    assert "grouped" not in report.skipped[0].reason


def test_synthetic_multi_option_boxes_expanded_and_restyled():
    """Tests slides with 1, 2, or 3 boxes containing 4 options (Slide 20, 26, 27/31 cases)."""
    prs = pptx.Presentation()

    # Slide 1: 1 question + 1 box containing all 4 options
    s1 = prs.slides.add_slide(prs.slide_layouts[6])
    lbl1 = s1.shapes.add_textbox(Inches(1), Inches(0.5), Inches(3), Inches(0.5))
    lbl1.text = "QUESTION 20"
    q1 = s1.shapes.add_textbox(Inches(1), Inches(1.2), Inches(8), Inches(1))
    q1.text = "What is 20% of 250?"
    opts1 = s1.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(3))
    opts1.text_frame.text = "A. 50\nB. 40\nC. 60\nD. 45"

    # Slide 2: 1 question + 2 boxes with 2 options each (2 columns)
    s2 = prs.slides.add_slide(prs.slide_layouts[6])
    lbl2 = s2.shapes.add_textbox(Inches(1), Inches(0.5), Inches(3), Inches(0.5))
    lbl2.text = "QUESTION 26"
    q2 = s2.shapes.add_textbox(Inches(1), Inches(1.2), Inches(8), Inches(1))
    q2.text = "Express 3/4 as a percentage:"
    col1 = s2.shapes.add_textbox(Inches(1), Inches(2.5), Inches(4), Inches(2))
    col1.text_frame.text = "A. 75%\nB. 25%"
    col2 = s2.shapes.add_textbox(Inches(5), Inches(2.5), Inches(4), Inches(2))
    col2.text_frame.text = "C. 50%\nD. 80%"

    # Slide 3: 1 question + 1 box with 2 options + 2 boxes with 1 option each
    s3 = prs.slides.add_slide(prs.slide_layouts[6])
    lbl3 = s3.shapes.add_textbox(Inches(1), Inches(0.5), Inches(3), Inches(0.5))
    lbl3.text = "QUESTION 27"
    q3 = s3.shapes.add_textbox(Inches(1), Inches(1.2), Inches(8), Inches(1))
    q3.text = "If price increases by 10% and then decreases by 10%:"
    box_ab = s3.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(1.5))
    box_ab.text_frame.text = "A. Price remains same\nB. Price decreases by 1%"
    box_c = s3.shapes.add_textbox(Inches(1), Inches(4.2), Inches(8), Inches(0.8))
    box_c.text_frame.text = "C. Price increases by 1%"
    box_d = s3.shapes.add_textbox(Inches(1), Inches(5.2), Inches(8), Inches(0.8))
    box_d.text_frame.text = "D. Cannot be determined"

    buf = io.BytesIO()
    prs.save(buf)

    out_bytes, report = process(buf.getvalue())
    assert report.total_slides == 3
    # All 3 slides must be detected as polls!
    assert report.poll_slides == [1, 2, 3]
    assert report.poll_count == 3
    assert report.skipped == []



@needs_discount
def test_discount_deck_detects_all_quiz_slides_and_slide_25():
    prs = pptx.Presentation(str(DISCOUNT_DECK))
    detection = find_polls(prs)

    poll_nums = [p.number for p in detection.polls]
    assert poll_nums == [10, 11, 12, 15, 16, 18, 22, 23, 25, 26]

    p25 = next(p for p in detection.polls if p.number == 25)
    assert "MCQs" in p25.label.text_frame.text
    assert "A customer pays ₹1,180" in p25.question.text_frame.text
    assert [o.text_frame.text.strip() for o in p25.options] == [
        "₹900",
        "₹1,000",
        "₹1,050",
        "₹1,100",
    ]


@needs_discount
def test_discount_deck_pipeline_runs_and_preserves():
    with open(DISCOUNT_DECK, "rb") as f:
        in_bytes = f.read()

    out_bytes, report = process(in_bytes)
    assert report.total_slides == 27
    assert report.poll_slides == [10, 11, 12, 15, 16, 18, 22, 23, 25, 26]

