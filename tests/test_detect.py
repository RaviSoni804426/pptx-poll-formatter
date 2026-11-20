import os
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Emu

from poll_formatter.detect import find_polls

# Real decks aren't committed (size + privacy). Point at a local copy.
_LOCAL_TEST_DECK = Path("test_decks") / "Into the Night Sky!_06_AS_CBSE revamped.pptx"
_DOWNLOADS_DECK = Path.home() / "Downloads" / "Into the Night Sky!_06_AS_CBSE revamped.pptx"
_DEFAULT_DECK = _LOCAL_TEST_DECK if _LOCAL_TEST_DECK.exists() else _DOWNLOADS_DECK

NIGHT_SKY = Path(os.environ.get("NIGHT_SKY_DECK", _DEFAULT_DECK))
needs_deck = pytest.mark.skipif(not NIGHT_SKY.exists(), reason=f"missing {NIGHT_SKY}")


@pytest.fixture(scope="module")
def night_sky():
    return Presentation(str(NIGHT_SKY))


@needs_deck
def test_finds_exactly_the_quiz_slides(night_sky):
    got = [p.number for p in find_polls(night_sky).polls]
    assert got == [8, 9, 12, 19, 20, 21, 24, 25, 26, 27]


@needs_deck
def test_lookalikes_are_not_polls(night_sky):
    # 3, 22, 23: heading + 4 short boxes.  28: "REFLECTION QUESTIONS" + 3 prompts.
    got = {p.number for p in find_polls(night_sky).polls}
    assert not got & {3, 5, 22, 23, 28}


@needs_deck
def test_reads_question_and_options_in_grid_order(night_sky):
    p = next(p for p in find_polls(night_sky).polls if p.number == 8)
    assert p.label.text_frame.text.strip() == "QUESTION"
    assert p.question.text_frame.text.startswith("Why can villagers usually see more stars")
    assert [o.text_frame.text for o in p.options] == [
        "Villages have more open land",        # row 1 left  -> A
        "Villages have fewer buildings",       # row 1 right -> B
        "Villages have less light pollution",  # row 2 left  -> C
        "Villages are closer to stars",        # row 2 right -> D
    ]
    assert [e.text_frame.text for e in p.extras] == ["Concept Check"]


def _box(slide, text, x, y, w=3_000_000, h=500_000):
    tb = slide.shapes.add_textbox(Emu(x), Emu(y), Emu(w), Emu(h))
    tb.text_frame.text = text
    return tb


def test_near_miss_is_reported_not_silently_changed():
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])  # blank
    _box(s, "QUIZ", 100_000, 100_000)
    _box(s, "Which planet is red?", 100_000, 800_000, w=8_000_000)
    for i, t in enumerate(["Mars", "Venus", "Earth"]):  # only 3 options
        _box(s, t, 100_000 + i * 3_100_000, 2_000_000)
    d = find_polls(prs)
    assert d.polls == []
    assert [k.number for k in d.skipped] == [1]
    assert "3" in d.skipped[0].reason


def test_vertical_list_of_four_options():
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    _box(s, "Pick the prime number:", 100_000, 100_000, w=8_000_000)
    for i, t in enumerate(["4", "6", "7", "9"]):  # 4 rows, no label, question ends in ':'
        _box(s, t, 100_000, 1_000_000 + i * 700_000)
    (p,) = find_polls(prs).polls
    assert p.label is None
    assert [o.text_frame.text for o in p.options] == ["4", "6", "7", "9"]
