# Journey-time monthly file builder - upload CSV, download all Excel files as a ZIP.
# Run:  streamlit run app.py      (keep journeytime_tool.py in the same folder)
import pandas as pd
import streamlit as st
import journeytime_tool as jt

st.set_page_config(page_title="Journey Time File Builder", page_icon="🚦", layout="centered")
st.title("🚦 Journey Time File Builder")
st.write("Upload the Google travel-time CSV exported from the database. "
         "You get one Excel file per month (one day each, 06:00–21:00, 10-minute rows) in a ZIP.")

up = st.file_uploader("1) Google travel-time CSV (the new data)", type="csv")
bf_file = st.file_uploader("2) Before file (optional) - the fixed 'before' day, added to every month",
                           type=["csv", "xlsx"],
                           help="Needs a Time column and a column like '11-09-2024 (Before)'. Any 'After' column in it is ignored.")
with st.expander("Settings", expanded=False):
    wd = st.selectbox("Preferred weekday", jt.WD, index=3)
    min_cov = st.slider("Minimum real data needed to use a day", 0.05, 1.0, 0.20, 0.05,
                        help="Share of the 90 time slots that must have real readings")
    fill_months = st.checkbox("Create a predicted day for months with no data", True)
    force_wd = st.checkbox("If a month has no preferred-weekday data, predict that weekday instead of using another real day", False)
    hl = st.checkbox("Highlight predicted cells in yellow", True)
    hdr = st.selectbox("Header cells B1 / C1", ["Real date cells (default)", "Text: date + weekday", "Text: date only"],
                       help="Use the first one unless the portal rejects the file.")
    hdr_mode = {"Real date cells (default)": "date", "Text: date + weekday": "weekday", "Text: date only": "text"}[hdr]
    natural = st.checkbox("Add natural variation to predicted cells (otherwise they look too smooth)", True)

if up:
    try:
        df = jt.load(up)
        before = jt.load_before(bf_file) if bf_file else None
        res = jt.process(df, wd, min_cov, fill_months, force_wd, before, natural)
    except Exception as e:
        st.error(f"Could not process the file: {e}"); st.stop()
    if not res:
        st.warning("No usable days found."); st.stop()
    log = jt.log_frame(res)
    st.success(f"{len(res)} monthly files ready.")
    if before:
        nb = int(log["before_predicted_slots"].iloc[0])
        st.info(f"Before day {before[0]:%d-%m-%Y}: {90 - nb} real time slots, {nb} predicted. "
                f"Every file name starts with {before[0]:%Y-%m-%d}.")
    st.dataframe(log[["month", "date", "weekday", "source", "predicted_slots"]], use_container_width=True, hide_index=True)
    st.download_button("⬇️ Download all files (ZIP)", jt.build_zip(res, hl, hdr_mode),
                       file_name="journeytime_monthly_files.zip", mime="application/zip", type="primary")
    pick = st.selectbox("Preview a month", [r["month"] for r in res])
    r = next(x for x in res if x["month"] == pick)
    cols = {"After": r["values"]}
    if r.get("before"):
        cols = {"Before": r["before"][1], **cols}
    st.line_chart(pd.DataFrame(cols, index=[t.strftime("%H:%M") for t in jt.SLOTS]))
    st.caption("Yellow cells in the Excel files (and 'predicted_slots' above) are estimates, not Google readings.")
