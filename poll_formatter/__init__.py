"""PPTX Poll Formatter package."""
from .detect import Detection, PollSlide, Skipped, find_polls
from .restyle import restyle
from .rebuild import build_polls_deck
from .pipeline import ProcessReport, process

__all__ = [
    "Detection",
    "PollSlide",
    "Skipped",
    "find_polls",
    "restyle",
    "build_polls_deck",
    "ProcessReport",
    "process",
]
