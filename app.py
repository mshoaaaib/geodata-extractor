"""
GeoData Extractor — Streamlit Web UI

Upload an Excel file with flood event data → extract 13 geospatial features →
download the enriched Excel file.
"""

import io
import os
import threading
import queue

import pandas as pd
import streamlit as st

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="GeoData Extractor",
    page_icon="🌍",
    layout="wide",
)

# ── Inject API credentials from Streamlit secrets / env vars ─────────────────
def _apply_credentials():
    """
    Credentials come from st.secrets (Streamlit Cloud) or environment variables.
    Streamlit Cloud: add them in Settings → Secrets as:
        SH_CLIENT_ID     = "..."
        SH_CLIENT_SECRET = "..."
        GEE_PROJECT      = "..."
    """
    for key in ("SH_CLIENT_ID", "SH_CLIENT_SECRET", "GEE_PROJECT"):
        try:
            val = st.secrets.get(key) or os.environ.get(key)
            if val:
                os.environ[key] = val
        except Exception:
            pass

_apply_credentials()

# ── Import the extractor (after env vars are set) ────────────────────────────
try:
    import extract_geodata as geo
    _IMPORT_OK = True
    _IMPORT_ERR = ""
except Exception as e:
    _IMPORT_OK = False
    _IMPORT_ERR = str(e)


# ─────────────────────────────────────────────────────────────────────────────
# UI
# ─────────────────────────────────────────────────────────────────────────────

st.title("🌍 GeoData Extractor")
st.caption(
    "Upload an Excel file with flood event coordinates and dates. "
    "The tool extracts 13 geospatial features per row and returns "
    "an enriched Excel file."
)

if not _IMPORT_OK:
    st.error(f"Failed to load extractor module: {_IMPORT_ERR}")
    st.stop()

# ── Sidebar — credentials override ───────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")

    st.subheader("Sentinel Hub")
    sh_id = st.text_input(
        "Client ID",
        value=os.environ.get("SH_CLIENT_ID", geo.SENTINELHUB_CLIENT_ID),
        type="password",
    )
    sh_secret = st.text_input(
        "Client Secret",
        value=os.environ.get("SH_CLIENT_SECRET", geo.SENTINELHUB_CLIENT_SECRET),
        type="password",
    )

    st.subheader("Google Earth Engine")
    gee_project = st.text_input(
        "GEE Project ID",
        value=os.environ.get("GEE_PROJECT", "fleet-furnace-348411"),
    )

    if st.button("Apply credentials"):
        geo.SENTINELHUB_CLIENT_ID     = sh_id
        geo.SENTINELHUB_CLIENT_SECRET = sh_secret
        try:
            import ee
            ee.Initialize(project=gee_project)
            geo._GEE_AVAILABLE = True
            st.success("Credentials applied ✓")
        except Exception as e:
            st.warning(f"GEE init failed (NDVI/population will use fallbacks): {e}")

    st.markdown("---")
    st.subheader("Expected columns")
    st.markdown(
        "| Column | Example |\n"
        "|---|---|\n"
        "| `lat` / `lat_approx` | 33.597 |\n"
        "| `lon` / `lot_approx` | 73.045 |\n"
        "| `flood_date` | 2022-08-25 |\n"
        "| `end_date` | 2022-08-29 |\n"
        "\n"
        "Or use **Year / Month / Day** columns with day ranges like `24–27`.\n\n"
        "A **location / area** column can substitute for coordinates."
    )

# ── File upload ───────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Upload input Excel file (.xlsx)",
    type=["xlsx"],
    help="Must contain lat/lon (or location name) and flood date columns.",
)

if uploaded is None:
    st.info("Upload an Excel file to get started.")
    st.stop()

df_input = pd.read_excel(uploaded)
st.subheader(f"Preview — {len(df_input)} rows × {len(df_input.columns)} columns")
st.dataframe(df_input, use_container_width=True)

# ── Run extraction ────────────────────────────────────────────────────────────
if st.button("▶ Run Extraction", type="primary", use_container_width=True):

    progress_bar  = st.progress(0.0, text="Starting…")
    log_container = st.expander("Live log", expanded=True)
    log_area      = log_container.empty()
    log_lines: list[str] = []

    result_holder: dict = {}
    error_holder:  dict = {}
    log_queue: queue.Queue = queue.Queue()

    def _log(msg: str):
        log_queue.put(msg)

    def _progress(current: int, total: int):
        log_queue.put(f"__PROGRESS__{current}/{total}")

    def _run():
        try:
            df_out, failed = geo.process_dataframe(
                df_input,
                log_fn=_log,
                progress_fn=_progress,
            )
            result_holder["df"]     = df_out
            result_holder["failed"] = failed
        except Exception as e:
            error_holder["msg"] = str(e)
        finally:
            log_queue.put("__DONE__")

    thread = threading.Thread(target=_run, daemon=True)
    thread.start()

    # Drain the queue and update UI while the thread runs
    while True:
        try:
            msg = log_queue.get(timeout=0.5)
        except queue.Empty:
            if not thread.is_alive():
                break
            continue

        if msg == "__DONE__":
            break
        elif msg.startswith("__PROGRESS__"):
            cur, tot = map(int, msg[len("__PROGRESS__"):].split("/"))
            frac = cur / tot if tot else 1.0
            progress_bar.progress(frac, text=f"Row {cur} / {tot}")
        else:
            log_lines.append(msg)
            log_area.text("\n".join(log_lines[-120:]))  # show last 120 lines

    thread.join()

    # ── Show result ───────────────────────────────────────────────────────────
    if error_holder:
        st.error(f"Extraction failed: {error_holder['msg']}")
        st.stop()

    progress_bar.progress(1.0, text="Done ✓")

    df_out    = result_holder["df"]
    failed    = result_holder["failed"]
    succeeded = len(df_input) - len(failed)

    col1, col2, col3 = st.columns(3)
    col1.metric("Total rows",    len(df_input))
    col2.metric("Extracted",     succeeded)
    col3.metric("Failed / skipped", len(failed))

    st.subheader("Results preview")
    st.dataframe(df_out, use_container_width=True)

    # ── Download button ───────────────────────────────────────────────────────
    buf = io.BytesIO()
    df_out.to_excel(buf, index=False)
    buf.seek(0)

    original_stem = os.path.splitext(uploaded.name)[0]
    download_name = f"{original_stem}_extracted.xlsx"

    st.download_button(
        label="⬇ Download extracted Excel",
        data=buf,
        file_name=download_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
