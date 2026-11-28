# PPTX Poll Formatter

A Streamlit app that finds poll/quiz slides in PowerPoint lecture decks and converts them into the exact slide format AI Tutor (Tuto LMS) imports, leaving every non-poll slide byte-identical.

---

## Output Modes

### 1. Tuto-Ready (default) — `mode="inplace"`

Replaces every detected poll slide **inside the original package** with the
23-shape blueprint AI Tutor is known to import. Everything else in the deck is
byte-identical.

Per poll slide: Blank layout, slide-level background `#FBF6EC`, and these shapes
in this exact order (AI Tutor reads the text in document order):

| # | Name | Type | Geometry (EMU) | Colour | Content |
|---|------|------|----------------|--------|---------|
| 0 | Text 0 | TextBox | 960144, 685817, 9601440, 480072 | `#2E6FBE` | `POLL N` — Calibri 18 bold, left/middle |
| 1 | Shape 1 | roundRect | 960144, 1303053, 16322448, 2057451 | `#E8F1FD` | question card |
| 2 | Shape 2 | ellipse | 15636631, 1645961, 1371634² | `#4A90E2` | icon circle |
| 3 | Image 0 | Picture | 15965823, 1975153, 713250² | — | 256×256 PNG |
| 4 | Text 3 | TextBox | 1440216, 1303053, 13922088, 2057451 | `#3B2A5E` | question — Calibri 28.5 bold, middle |
| 5/9/13/17 | Shape 4/8/12/16 | roundRect | 7886897 × 2126033 | `#FFFFFF` + shadow | option cards A–D |
| 6/10/14/18 | Shape 5/9/13/17 | ellipse | 960144² | `#4A90E2` | letter circles |
| 7/11/15/19 | Text 6/10/14/18 | TextBox | 960144² | `#FFFFFF` | `A`–`D` — Calibri 30 bold, centre/middle |
| 8/12/16/20 | Text 7/11/15/19 | TextBox | 5966609 × 2126033 | `#3B2A5E` | option text — Calibri 21.75 regular, middle |
| 21 | Text 20 | TextBox | 960144, 8984205, 16322448, 548654 | `#8B849A` | hint — Calibri 21 regular |
| 22 | Text 21 | TextBox | 17327856, 9711171, 548654, 411490 | `#8B849A` | `N` — Calibri 15 regular, right/middle |

All units EMU at `REF_WIDTH = 18288000`; other slide widths scale by
`slide_width / REF_WIDTH`. Defined in one place: `poll_formatter/spec.py`.

**Trade-off:** poll slides lose the source deck's Canva artwork. This is the only
mode verified to import reliably.

### 2. Restyled — keeps original artwork

In-place restyling that preserves backgrounds, freeforms and images. Produces the
correct font hierarchy and shape order but **has not imported reliably into AI
Tutor**. Use only if the artwork matters more than automatic poll detection.

### 3. Polls Only — standalone deck

Just the extracted polls on the blueprint, as a separate presentation.

---

## Rules

- **Detection**: a `Concept Check` / `Poll question` style label plus exactly
  1 question box and 4 option boxes. Anything else is reported as skipped, never
  silently altered.
- **Option order**: read from the source 2×2 grid, row first (top→bottom) then
  column (left→right): top-left `A`, top-right `B`, bottom-left `C`, bottom-right `D`.
- **Text**: question and option wording is preserved; only surrounding whitespace
  is trimmed and any leading `A.` / `1)` marker removed.
- **Numbering**: `POLL N` and the corner number both start at 1 in deck order.
- **Scope**: work happens inside the uploaded package. Non-poll slides, layouts,
  masters and media are never touched.

---

## Local Development

### 1. Setup Environment

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
pip install pytest
```

### 2. Run Tests

```bash
pytest -v
```

Tests verify:
- Exact poll slide detection and grid-order reading on sample decks.
- Text restyling, font sizes, weights, and run colors.
- Letter spacing and option nudging.
- Byte-for-byte XML preservation of non-poll slides and non-text shapes.
- Pipeline and error handling.

### 3. Run Streamlit App Locally

```bash
streamlit run app.py
```

Open `http://localhost:8501` in your browser.

---

## Deploying to Streamlit Cloud

1. **Push to GitHub**:
   ```bash
   git add .
   git commit -m "Initial commit of pptx-poll-formatter"
   git push origin main
   ```
2. **Connect Streamlit Cloud**:
   - Go to [share.streamlit.io](https://share.streamlit.io).
   - Click **New app**.
   - Select your repository and branch (`main`).
   - Set **Main file path** to `app.py`.
   - Click **Deploy!**.
3. **Upload Limit**:
   - `.streamlit/config.toml` configures `maxUploadSize = 200` (200 MB), accommodating large lecture presentations with embedded imagery.
# pptx-poll-formatter
