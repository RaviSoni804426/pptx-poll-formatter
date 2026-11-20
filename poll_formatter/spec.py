"""The poll format AI Tutor accepts, measured from slide 1 of the accepted deck
('All Polls - Density, Discount, Profit and Loss.pptx', 18288000 EMU wide).

Every size here is for a slide REF_WIDTH wide; multiply by (slide_width / REF_WIDTH)
for other slide sizes. Change the format in this one file only.
"""

REF_WIDTH = 18288000  # EMU

FONT = "Calibri"

# (point size, bold) per element
LABEL = (18.0, True)      # "POLL N"
QUESTION = (28.5, True)
LETTER = (30.0, True)     # "A" / "B" / "C" / "D"
OPTION = (21.75, False)

LETTERS = ("A", "B", "C", "D")
LETTER_BOX = 960144       # EMU, square letter box on the reference slide

# Corner / chip labels seen on source quiz slides. Matched case-insensitively.
# A slide carrying one of these is a poll candidate.
LABEL_WORDS = {
    "question",
    "questions",
    "quick quiz",
    "mcqs",
    "mcq",
    "mcq's",
    "quiz",
    "quizzes",
    "concept check",
    "poll",
    "polls",
    "check your understanding",
    "multiple choice questions",
    "multiple choice question",
    "practice question",
    "practice questions",
    "test your knowledge",
    "try it yourself",
}
