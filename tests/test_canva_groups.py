import os
from pathlib import Path
import pytest
import pptx
from pptx.util import Emu

from poll_formatter.detect import find_polls
from poll_formatter.pipeline import process

DISCOUNT_DECK = Path("test_decks") / "Discount in Shopping.pptx"
needs_discount = pytest.mark.skipif(not DISCOUNT_DECK.exists(), reason=f"missing {DISCOUNT_DECK}")


@needs_discount
def test_discount_deck_detects_all_quiz_slides_and_slide_25():
    prs = pptx.Presentation(str(DISCOUNT_DECK))
    detection = find_polls(prs)

    poll_nums = [p.number for p in detection.polls]
    # Exactly the 10 quiz slides, including Slide 25 (the user's screenshot)
    assert poll_nums == [10, 11, 12, 15, 16, 18, 22, 23, 25, 26]

    # Verify Slide 25 specifically
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
    # Slide 4 has genuinely grouped math equations and is safely skipped
    assert any(s.number == 4 for s in report.skipped)
