"""Pipeline for processing presentations: detect, rebuild, restyle, and preserve."""
from dataclasses import dataclass
import io
import zipfile
import pptx

from . import spec
from .detect import Detection, PollSlide, Skipped, find_polls
from .restyle import restyle
from .rebuild import build_polls_deck, build_inplace_deck


@dataclass
class ProcessReport:
    total_slides: int
    poll_slides: list[int]
    skipped: list[Skipped]
    non_poll_count: int
    ai_tutor_bytes: bytes = b""
    restyled_bytes: bytes = b""
    inplace_bytes: bytes = b""

    @property
    def poll_count(self) -> int:
        return len(self.poll_slides)


def process(pptx_bytes: bytes, mode: str = "restyle") -> tuple[bytes, ProcessReport]:
    """Process a PPTX file in memory:
    - Finds all poll slides across any layout/format
    - Generates proven AI Tutor rebuilt deck (polls only)
    - Generates in-place restyled deck (non-poll slides preserved)
    Returns (primary_bytes, ProcessReport).
    If mode == 'rebuild', primary_bytes is the AI Tutor rebuilt deck.
    If mode == 'restyle', primary_bytes is the in-place restyled deck.
    """
    try:
        prs = pptx.Presentation(io.BytesIO(pptx_bytes))
    except Exception as e:
        raise ValueError(f"Could not open PPTX file: {e}") from e

    total_slides = len(prs.slides)
    scale = prs.slide_width / spec.REF_WIDTH

    detection = find_polls(prs)

    # 1. Build Proven AI Tutor Deck (polls only, standard 23-shape blueprint)
    ai_tutor_bytes = b""
    inplace_bytes = b""
    if detection.polls:
        ai_tutor_bytes = build_polls_deck(detection.polls)
        # Full deck, poll slides swapped to the proven blueprint. Built BEFORE
        # restyle() mutates the shapes it reads from.
        inplace_bytes = build_inplace_deck(pptx_bytes, detection.polls)

    # 2. Build In-Place Restyled Deck (preserving non-poll slides byte-identically)
    modified_parts: dict[str, bytes] = {}
    for poll_no, poll in enumerate(detection.polls, 1):
        slide = prs.slides[poll.number - 1]
        restyle(slide, poll, poll_no, scale)
        partname = slide.part.partname.lstrip("/")
        modified_parts[partname] = slide.part.blob

    out_buf = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(pptx_bytes), "r") as z_in:
        with zipfile.ZipFile(out_buf, "w", compression=zipfile.ZIP_DEFLATED) as z_out:
            for item in z_in.infolist():
                if item.filename in modified_parts:
                    z_out.writestr(item, modified_parts[item.filename])
                else:
                    z_out.writestr(item, z_in.read(item.filename))

    restyled_bytes = out_buf.getvalue()

    if mode == "rebuild" and ai_tutor_bytes:
        primary_bytes = ai_tutor_bytes
    elif mode == "inplace" and inplace_bytes:
        primary_bytes = inplace_bytes
    else:
        primary_bytes = restyled_bytes

    report = ProcessReport(
        total_slides=total_slides,
        poll_slides=[p.number for p in detection.polls],
        skipped=detection.skipped,
        non_poll_count=total_slides - len(detection.polls),
        ai_tutor_bytes=ai_tutor_bytes,
        restyled_bytes=restyled_bytes,
        inplace_bytes=inplace_bytes,
    )

    return primary_bytes, report
