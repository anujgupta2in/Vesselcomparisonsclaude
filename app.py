"""Streamlit app: add Job Description, Maker and Model from Job List to Job Status."""

from pathlib import Path

import pandas as pd
import streamlit as st

import merge_logic as ml

st.set_page_config(page_title="Job Status + Job List Merger", page_icon="⚓", layout="wide")
st.title("⚓ Job Status + Job List Merger")
st.caption(
    "Adds **Job Description, Maker and Model** from the Job List to every row of the Job Status file. "
    "Rows are matched on Job Code + Machinery + Component, so a Job Code used on several machineries "
    "gets the correct maker/model for each one."
)

c1, c2 = st.columns(2)
with c1:
    status_file = st.file_uploader("1. Job Status file", type=["csv", "xlsx", "xls"], key="status")
with c2:
    list_file = st.file_uploader("2. Job List file", type=["csv", "xlsx", "xls"], key="list")

if not (status_file and list_file):
    st.info("Upload both files to start.")
    st.stop()


@st.cache_data(show_spinner="Matching jobs...")
def run(status_bytes, status_name, list_bytes, list_name):
    import io
    js = ml.read_table(io.BytesIO(status_bytes), status_name)
    jl = ml.read_table(io.BytesIO(list_bytes), list_name)
    merged, summary, warns = ml.merge(js, jl)
    return merged, summary, warns, len(jl)


try:
    merged, summary, warns, n_list = run(
        status_file.getvalue(), status_file.name, list_file.getvalue(), list_file.name)
except ValueError as e:
    st.error(f"{e}. Check that the files are in the right upload boxes.")
    st.stop()

for w in warns:
    st.warning(w)

# ---- Summary ------------------------------------------------------------- #
total = len(merged)
review = summary[ml.LEVEL_3] + summary[ml.LEVEL_4]
m = st.columns(5)
m[0].metric("Job Status rows", f"{total:,}")
m[1].metric("Job List rows", f"{n_list:,}")
m[2].metric("Matched", f"{total - summary[ml.NOT_FOUND]:,}",
            f"{(total - summary[ml.NOT_FOUND]) / total:.1%}" if total else None, delta_color="off")
m[3].metric("Needs review", f"{review:,}")
m[4].metric("Not found", f"{summary[ml.NOT_FOUND]:,}")

with st.expander("Match level breakdown", expanded=review > 0 or summary[ml.NOT_FOUND] > 0):
    st.dataframe(
        pd.DataFrame({"Match Level": list(summary), "Rows": list(summary.values())}),
        hide_index=True, width="content")
    st.markdown(
        "- **1** – Job Code, Machinery and Component all match (most reliable)\n"
        "- **2** – Job Code and Machinery match; only one Job List entry exists for them\n"
        "- **3** – Several Job List entries for the same Job Code/Machinery with different "
        "maker/model; all values shown separated by ` | ` – **please review**\n"
        "- **4** – Machinery name differs, matched on Job Code alone (code used on one machinery only)\n"
        "- **Not Found** – no matching Job Code / Machinery in the Job List"
    )

# ---- Preview with filters ------------------------------------------------- #
st.subheader("Preview")
f1, f2, f3 = st.columns([2, 2, 3])
levels = f1.multiselect("Match Level", list(summary), default=[k for k, v in summary.items() if v])
code_col = next((c for c in merged.columns if str(c).lower() == "job code"), None)
mach_col = next((c for c in merged.columns if str(c).lower() in ("machinery location", "machinery")), None)
machs = f2.multiselect("Machinery", sorted(merged[mach_col].unique()) if mach_col else [])
search = f3.text_input("Search (Job Code, Title, Maker, Model...)")

view = merged[merged["Match Level"].isin(levels)]
if machs:
    view = view[view[mach_col].isin(machs)]
if search:
    s = search.lower()
    view = view[view.apply(lambda r: s in " ".join(map(str, r.values)).lower(), axis=1)]

show_cols = [c for c in ["Machinery Location", "Sub Component Location", "Job Code", "Title",
                          "Maker", "Model", "Verifying Rank", "Match Level", "Description", "Match Note"]
             if c in view.columns]
show_all = st.toggle("Show all columns", value=False)
st.caption(f"{len(view):,} of {total:,} rows")
st.dataframe(
    view if show_all else view[show_cols],
    hide_index=True, width="stretch", height=480,
    column_config={"Description": st.column_config.TextColumn(width="large")},
)

# ---- Downloads ------------------------------------------------------------ #
st.subheader("Download")
base = Path(status_file.name).stem + " - with Job Details"
d1, d2, _ = st.columns([1, 1, 3])


@st.cache_data(show_spinner="Building Excel file...")
def excel_bytes(df):
    return ml.to_excel_bytes(df)


d1.download_button("⬇ Excel (.xlsx)", excel_bytes(merged), f"{base}.xlsx",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   type="primary", width="stretch")
d2.download_button("⬇ CSV", ml.to_csv_bytes(merged), f"{base}.csv", "text/csv",
                   width="stretch")
st.caption("Full file in the standard column order (Vessel ... Attachment Indicator), with Description, "
           "Verifying Rank, Maker and Model from the Job List, plus Match Level and Match Note at the end. "
           "Rows needing review are highlighted in the Excel file.")
