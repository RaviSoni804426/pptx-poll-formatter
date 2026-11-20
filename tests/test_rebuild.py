import io
from pathlib import Path
import pytest
import pptx

from poll_formatter.detect import find_polls
from poll_formatter.rebuild import build_polls_deck
from poll_formatter.pipeline import process

NIGHT_SKY = Path("test_decks") / "Into the Night Sky!_06_AS_CBSE revamped.pptx"
DISCOUNT = Path("test_decks") / "Discount in Shopping.pptx"


@pytest.mark.skipif(not NIGHT_SKY.exists(), reason="missing Night Sky deck")
def test_rebuild_night_sky_deck():
    prs_in = pptx.Presentation(str(NIGHT_SKY))
    d = find_polls(prs_in)
    assert len(d.polls) == 10

    deck_bytes = build_polls_deck(d.polls)
    out_prs = pptx.Presentation(io.BytesIO(deck_bytes))

    # Must contain exactly 10 slides (polls only)
    assert len(out_prs.slides) == 10

    # Slide 10 corresponds to "Orion is best seen in India during:"
    s10 = out_prs.slides[9]
    assert len(s10.shapes) == 23

    # Shape 0: POLL 10
    assert s10.shapes[0].text_frame.text.strip() == "POLL 10"
    assert (s10.shapes[0].left, s10.shapes[0].top) == (960144, 685817)

    # Shape 4: Question
    assert "Orion is best seen in India during:" in s10.shapes[4].text_frame.text

    # Letters A, B, C, D in circle positions
    assert s10.shapes[7].text_frame.text.strip() == "A"
    assert s10.shapes[11].text_frame.text.strip() == "B"
    assert s10.shapes[15].text_frame.text.strip() == "C"
    assert s10.shapes[19].text_frame.text.strip() == "D"

    # Options
    assert s10.shapes[8].text_frame.text.strip() == "May–July"
    assert s10.shapes[12].text_frame.text.strip() == "August–October"
    assert s10.shapes[16].text_frame.text.strip() == "December–April"
    assert s10.shapes[20].text_frame.text.strip() == "September–November"

    # Slide number
    assert s10.shapes[22].text_frame.text.strip() == "10"


@pytest.mark.skipif(not DISCOUNT.exists(), reason="missing Discount deck")
def test_rebuild_discount_deck():
    prs_in = pptx.Presentation(str(DISCOUNT))
    d = find_polls(prs_in)
    assert len(d.polls) == 10

    deck_bytes = build_polls_deck(d.polls)
    out_prs = pptx.Presentation(io.BytesIO(deck_bytes))
    assert len(out_prs.slides) == 10

    # Slide 9 is the GST question from the user's screenshot
    s9 = out_prs.slides[8]
    assert s9.shapes[0].text_frame.text.strip() == "POLL 9"
    assert "1,180" in s9.shapes[4].text_frame.text
    assert s9.shapes[8].text_frame.text.strip() == "₹900"
    assert s9.shapes[12].text_frame.text.strip() == "₹1,000"
    assert s9.shapes[16].text_frame.text.strip() == "₹1,050"
    assert s9.shapes[20].text_frame.text.strip() == "₹1,100"
