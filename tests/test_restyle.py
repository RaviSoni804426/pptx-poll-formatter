import os
from pathlib import Path
import pytest
import pptx
from pptx.util import Pt, Emu

from poll_formatter import spec
from poll_formatter.detect import find_polls
from poll_formatter.restyle import restyle

_LOCAL_TEST_DECK = Path("test_decks") / "Into the Night Sky!_06_AS_CBSE revamped.pptx"
_DOWNLOADS_DECK = Path.home() / "Downloads" / "Into the Night Sky!_06_AS_CBSE revamped.pptx"
_DEFAULT_DECK = _LOCAL_TEST_DECK if _LOCAL_TEST_DECK.exists() else _DOWNLOADS_DECK

NIGHT_SKY = Path(os.environ.get("NIGHT_SKY_DECK", _DEFAULT_DECK))
needs_deck = pytest.mark.skipif(not NIGHT_SKY.exists(), reason=f"missing {NIGHT_SKY}")


@pytest.fixture
def night_sky_deck():
    return pptx.Presentation(str(NIGHT_SKY))


@needs_deck
def test_restyle_slide_8_labels_and_fonts(night_sky_deck):
    scale = night_sky_deck.slide_width / spec.REF_WIDTH
    detection = find_polls(night_sky_deck)
    p8 = next(p for p in detection.polls if p.number == 8)
    slide8 = night_sky_deck.slides[7]

    restyle(slide8, p8, poll_no=1, scale=scale)

    # 1. Label text and font
    assert p8.label.text_frame.text.strip() == "POLL 1"
    run_label = p8.label.text_frame.paragraphs[0].runs[0]
    assert run_label.font.name == "Calibri"
    assert run_label.font.size == Pt(18.0 * scale)
    assert run_label.font.bold is True
    # Label color preserved (FDE08F)
    assert run_label.font.color.rgb == pptx.dml.color.RGBColor(0xFD, 0xE0, 0x8F)

    # 2. Question font and text preserved
    assert p8.question.text_frame.text.strip() == "Why can villagers usually see more stars than city people?"
    for p in p8.question.text_frame.paragraphs:
        for r in p.runs:
            assert r.font.name == "Calibri"
            assert r.font.size == Pt(28.5 * scale)
            assert r.font.bold is True

    # 3. Options fonts and text preserved
    expected_texts = [
        "Villages have more open land",
        "Villages have fewer buildings",
        "Villages have less light pollution",
        "Villages are closer to stars",
    ]
    for opt, expected_txt in zip(p8.options, expected_texts):
        assert opt.text_frame.text.strip() == expected_txt
        for p in opt.text_frame.paragraphs:
            for r in p.runs:
                assert r.font.name == "Calibri"
                assert r.font.size == Pt(21.75 * scale)
                assert r.font.bold is None

    # 4. Shape order at the end of shape tree
    # Trailing shapes must be: POLL 1, Question, A, OptA, B, OptB, C, OptC, D, OptD, Concept Check
    all_shapes = list(slide8.shapes)
    trailing = [s.text_frame.text.strip() for s in all_shapes[-13:]]
    assert trailing == [
        "Concept Check",
        "POLL 1",
        "Why can villagers usually see more stars than city people?",
        "A",
        "Villages have more open land",
        "B",
        "Villages have fewer buildings",
        "C",
        "Villages have less light pollution",
        "D",
        "Villages are closer to stars",
        "Think it through \u2014 which one is correct?",
        "1",
    ]


@needs_deck
def test_letters_do_not_overlap_options(night_sky_deck):
    scale = night_sky_deck.slide_width / spec.REF_WIDTH
    detection = find_polls(night_sky_deck)
    p8 = next(p for p in detection.polls if p.number == 8)
    slide8 = night_sky_deck.slides[7]

    orig_opt_lefts = [opt.left for opt in p8.options]
    restyle(slide8, p8, poll_no=1, scale=scale)

    all_shapes = list(slide8.shapes)
    # The letters are shapes at all_shapes[-9], [-7], [-5], [-3]
    # Corresponding to A, B, C, D
    letter_shapes = [all_shapes[-11], all_shapes[-9], all_shapes[-7], all_shapes[-5]]
    for idx, (let_sh, opt_sh) in enumerate(zip(letter_shapes, p8.options)):
        assert let_sh.text_frame.text.strip() == spec.LETTERS[idx]
        assert let_sh.left == orig_opt_lefts[idx]
        # Option text box was nudged right
        assert opt_sh.left > let_sh.left
        # Gap between right edge of letter box (or left + letter size) and opt
        assert opt_sh.left >= let_sh.left + int(spec.LETTER_BOX * scale)


def test_restyle_slide_with_no_label():
    prs = pptx.Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    tb_q = s.shapes.add_textbox(Emu(100_000), Emu(100_000), Emu(8_000_000), Emu(500_000))
    tb_q.text_frame.text = "Pick the prime number:"
    opts = []
    for i, t in enumerate(["4", "6", "7", "9"]):
        tb = s.shapes.add_textbox(Emu(100_000), Emu(1_000_000 + i * 700_000), Emu(3_000_000), Emu(500_000))
        tb.text_frame.text = t
        opts.append(tb)

    detection = find_polls(prs)
    assert len(detection.polls) == 1
    p = detection.polls[0]
    assert p.label is None

    restyle(s, p, poll_no=3, scale=1.0)

    # Added label has text "POLL 3"
    all_texts = [shape.text_frame.text.strip() for shape in s.shapes]
    assert all_texts == [
        "POLL 3",
        "Pick the prime number:",
        "A",
        "4",
        "B",
        "6",
        "C",
        "7",
        "D",
        "9",
        "Think it through \u2014 which one is correct?",
        "3",
    ]
