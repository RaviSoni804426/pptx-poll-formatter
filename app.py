"""Streamlit UI for PPTX Poll Formatter.

Upload any presentation deck, automatically identify all poll questions,
convert them into the proven AI Tutor format, and download the ready deck.
"""
from pathlib import Path
import time
import streamlit as st

from poll_formatter.pipeline import process

st.set_page_config(
    page_title="PPTX Poll Formatter",
    page_icon="📊",
    layout="centered",
)

st.title("PPTX Poll Formatter")
st.caption("Universal tool to identify poll questions in any PPTX presentation and format them for AI Tutor.")

# Sidebar info
with st.sidebar:
    st.header("How It Works")
    st.markdown(
        """
        - **Universal Detection**: Works with Canva, PowerPoint, Google Slides, and custom templates.
        - **Automatic Extraction**: Extracts Question and Options (A, B, C, D) in grid or list order.
        - **In-Place Restyling**: Restyles poll slides into the format required for AI Tutor, keeping all lecture slides byte-identical.
        """
    )
    st.divider()
    st.caption("Max file upload size: 200 MB")

uploaded_file = st.file_uploader(
    "Choose a PowerPoint presentation (.pptx)",
    type=["pptx"],
    help="Upload your lecture or quiz deck. Files are processed in memory and never stored on disk.",
)

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    file_size_mb = len(file_bytes) / (1024 * 1024)

    st.write(f"📁 **File:** `{uploaded_file.name}` ({file_size_mb:.1f} MB)")

    if file_size_mb > 200:
        st.error(f"The uploaded file is {file_size_mb:.1f} MB, which exceeds the 200 MB limit.")
        st.stop()

    if st.button("Process & Extract Polls", type="primary"):
        progress_bar = st.progress(10, text="Analyzing presentation structure...")
        status_box = st.empty()

        try:
            time.sleep(0.1)
            progress_bar.progress(35, text="Detecting poll slides across layouts and templates...")
            time.sleep(0.1)

            progress_bar.progress(70, text="Restyling poll slides...")
            out_bytes, report = process(file_bytes, mode="inplace")

            progress_bar.progress(100, text="Complete!")
            time.sleep(0.1)
            progress_bar.empty()

            st.success("Polls processed and converted successfully!")

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Slides", report.total_slides)
            col2.metric("Polls Detected", report.poll_count)
            col3.metric("Non-Poll Slides", report.non_poll_count)
            col4.metric("Skipped", len(report.skipped))

            base_stem = Path(uploaded_file.name).stem

            MIME = ("application/vnd.openxmlformats-officedocument"
                    ".presentationml.presentation")

            if report.poll_count > 0:
                st.subheader("Download")
                st.caption(
                    "Use **Tuto-Ready** if AI Tutor does not detect the polls. "
                    "Its poll slides are byte-identical to the format AI Tutor "
                    "is known to import; only the poll slides lose the source "
                    "deck's custom artwork."
                )
                st.download_button(
                    label="⬇️ Full Deck — Tuto-Ready (proven poll layout)",
                    data=report.inplace_bytes,
                    file_name=f"{base_stem} - Tuto Ready.pptx",
                    mime=MIME,
                    type="primary",
                    use_container_width=True,
                )
                st.download_button(
                    label="⬇️ Full Deck — Restyled (keeps original artwork)",
                    data=report.restyled_bytes,
                    file_name=f"{base_stem} - Restyled Full.pptx",
                    mime=MIME,
                    use_container_width=True,
                )
                st.download_button(
                    label="⬇️ Polls Only (standalone deck)",
                    data=report.ai_tutor_bytes,
                    file_name=f"{base_stem} - Polls Only.pptx",
                    mime=MIME,
                    use_container_width=True,
                )

            # Details
            with st.expander("Detection & Poll Summary", expanded=True):
                if report.poll_slides:
                    st.markdown(
                        f"**Detected Poll Slides:** {', '.join(str(s) for s in report.poll_slides)}"
                    )
                else:
                    st.info("No 4-option poll slides detected in this deck.")

                if report.skipped:
                    st.warning("⚠️ **Near-Miss Slides (Skipped):**")
                    for s in report.skipped:
                        st.markdown(f"- Slide **{s.number}**: {s.reason}")
                else:
                    st.markdown("✅ No ambiguous or near-miss slides skipped.")

        except ValueError as e:
            progress_bar.empty()
            st.error(f"Could not process presentation: {e}")
        except Exception as e:
            progress_bar.empty()
            st.error(
                f"An unexpected error occurred while processing the presentation: {e}. "
                "Please verify the deck is a standard PPTX file."
            )
