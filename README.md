# PPTX Poll Formatter (Restyle Mode)

A Streamlit web application that finds poll/quiz slides in PowerPoint lecture decks, restyles their text into the exact font hierarchy and shape ordering required by AI Tutor, and preserves everything else (backgrounds, cards, freeforms, images, non-poll slides) byte-identically.

---

## Features

- **In-Place Restyling**: Preserves original slide backgrounds, freeforms, card shapes, and graphics.
- **AI Tutor Compatible Format**:
  - `POLL N` label: Calibri 18 pt, Bold (preserves original label color)
  - Question: Calibri 28.5 pt, Bold (preserves run colors)
  - Letters `A`–`D`: Calibri 30 pt, Bold (adds individual letter textboxes matching option text color)
  - Options: Calibri 21.75 pt, Regular (nudged to the right so letters do not overlap)
  - Shape order: Reorders text shapes to `[label, question, A, optA, B, optB, C, optC, D, optD, extras...]` at the end of the slide's shape tree.
- **Byte-Identical Preservation**: Non-poll slide XML parts and non-text shapes on poll slides are 100% byte-identical between input and output.
- **Smart Detection**: Detects slides with 1 question + 4 options in 2×2 grid or 4-row layout. Near-misses (e.g. 3 options) are reported as skipped rather than altered silently.
- **Zero Disk Footprint**: All files are processed in-memory per request; no user decks are stored on disk.

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
