"""Sister Vessels Job Description Comparison - Streamlit tool.

Run:  streamlit run app.py
"""
import io
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st

from compare_engine import read_extract, compare, build_excel, diff_html, SHEET_ORDER

st.set_page_config(page_title="Sister Vessel JD Comparison", page_icon="⚓", layout="wide")

COLOR_A, COLOR_B = "#1F77B4", "#FF7F0E"

st.markdown("""
<style>
div[data-testid="stMetricValue"] {font-size: 1.6rem;}
.badge {display:inline-block;padding:2px 10px;border-radius:12px;color:white;font-weight:600;font-size:0.85rem;margin-right:6px}
.hint {color:#666;font-size:0.92rem;margin-top:-6px;margin-bottom:10px}
</style>
""", unsafe_allow_html=True)


@st.cache_data(show_spinner="Reading extracts and comparing… (takes ~15 seconds)")
def run_compare(bytes_a, bytes_b):
    da, db = read_extract(io.BytesIO(bytes_a)), read_extract(io.BytesIO(bytes_b))
    return compare(da, db), da, db


ss = st.session_state
ss.setdefault("nav", "summary")


def go(page, **preset):
    """Button callback: switch page and optionally preset widgets on that page."""
    ss.nav = page
    for k, v in preset.items():
        ss[k] = v


def clear_filters():
    ss.f_func, ss.f_mach, ss.f_search = [], [], ""


def reset_files():
    for k in ["files", "xlsx", "f_func", "f_mach", "f_search"]:
        ss.pop(k, None)
    ss.nav = "summary"


# ====================================================================
# Step 1 - load files
# ====================================================================
if "files" not in ss:
    st.title("⚓ Sister Vessel Job Description Comparison")
    st.markdown("Compare the PMS job lists of two sister vessels and find every discrepancy: "
                "missing jobs, different frequencies, ranks, descriptions, critical flags and more.")
    st.subheader("Step 1 – Select the two Job Status extracts")
    c1, c2 = st.columns(2)
    picked = {}
    for col, tag, color in [(c1, "A", COLOR_A), (c2, "B", COLOR_B)]:
        with col, st.container(border=True):
            st.markdown(f"<span class='badge' style='background:{color}'>Vessel {tag}</span>", unsafe_allow_html=True)
            up = st.file_uploader(f"Upload extract for vessel {tag}", type=["xlsx"], key=f"up_{tag}")
            path = st.text_input(f"…or paste the file path for vessel {tag}", key=f"path_{tag}",
                                 placeholder=r"C:\Users\...\Vessel - with Job Details.xlsx").strip().strip('"')
            if up:
                picked[tag] = (up.getvalue(), up.name)
                st.success(f"✔ {up.name}")
            elif path:
                p = Path(path)
                if p.is_file():
                    picked[tag] = (p.read_bytes(), p.name)
                    st.success(f"✔ {p.name}")
                else:
                    st.error("File not found – check the path.")

    ready = len(picked) == 2
    st.subheader("Step 2 – Compare")
    if st.button("🔍 Compare vessels", type="primary", disabled=not ready, use_container_width=True):
        ss.files = (picked["A"][0], picked["B"][0])
        ss.nav = "summary"
        st.rerun()
    if not ready:
        st.caption("Select both files to enable the comparison.")

    with st.expander("ℹ️ What does the tool check?"):
        st.markdown("""
- **Counts** – jobs per Function, Machinery Location and Sub-component.
- **Missing jobs** – matched on *Machinery + Sub-component + Job Code*, with the likely reason.
- **Job details** – Frequency, Performing Rank, Description, Title, Verifying Rank, Critical flag, Job Source, E-Form.
- **Equipment** – Maker / Model differences.
- **Data quality** – sub-component names that repeat the machinery name as a prefix.
- **Export** – everything as a multi-sheet Excel report, or any table as CSV.
""")
    st.stop()

try:
    R, raw_a, raw_b = run_compare(*ss.files)
except ValueError as e:
    st.error(f"Could not read the files: {e}")
    st.button("← Choose other files", on_click=reset_files)
    st.stop()

NA, NB = R["names"]
sa, sb = f" [{NA}]", f" [{NB}]"
fd = R["findings"].set_index("Check")["Count"]
other_counts = R["Other Field Diff"]["Field"].value_counts()

PAGES = {
    "summary": ("🏠 Summary", None),
    "counts": ("🔢 Count differences", int((R["Machinery Count"]["Status"] != "Match").sum())),
    "missing": ("❓ Missing jobs", len(R["Jobs Only A"]) + len(R["Jobs Only B"])),
    "frequency": ("⏱ Frequency", len(R["Frequency Diff"])),
    "rank": ("👷 Performing rank", len(R["PerfRank Diff"])),
    "description": ("📝 Description", len(R["Description Diff"])),
    "other": ("🏷 Critical & other fields", len(R["Other Field Diff"])),
    "equipment": ("⚙ Maker/Model & naming", len(R["Maker Model Diff"]) + len(R["Naming Issue"])),
    "drill": ("🔍 Machinery drill-down", None),
    "export": ("⬇ Export report", None),
}

# ====================================================================
# Sidebar - vessels, navigation, filters
# ====================================================================
with st.sidebar:
    st.markdown(f"<span class='badge' style='background:{COLOR_A}'>A</span> **{NA}**", unsafe_allow_html=True)
    st.markdown(f"<span class='badge' style='background:{COLOR_B}'>B</span> **{NB}**", unsafe_allow_html=True)
    st.button("🔄 Compare other files", on_click=reset_files, use_container_width=True)
    st.divider()
    st.radio("Go to", list(PAGES), key="nav", label_visibility="collapsed",
             format_func=lambda k: PAGES[k][0] + (f"  ({PAGES[k][1]:,})" if PAGES[k][1] is not None else ""))
    st.divider()

    ss.setdefault("f_func", [])
    ss.setdefault("f_mach", [])
    ss.setdefault("f_search", "")
    n_active = bool(ss.f_func) + bool(ss.f_mach) + bool(ss.f_search)
    with st.expander(f"🔎 Filters{f'  ({n_active} active)' if n_active else ''}", expanded=bool(n_active)):
        all_funcs = sorted(set(raw_a["Function"].dropna()) | set(raw_b["Function"].dropna()))
        st.multiselect("Function", all_funcs, key="f_func")
        mach_pool = R["Machinery Count"]
        if ss.f_func:
            mach_pool = mach_pool[mach_pool["Function"].isin(ss.f_func)]
        st.multiselect("Machinery Location", sorted(mach_pool["Machinery Location"].unique()), key="f_mach")
        st.text_input("Search (title, sub-component, job code)", key="f_search")
        st.button("✖ Clear filters", on_click=clear_filters, use_container_width=True, disabled=not n_active)

mach_in_func = set(mach_pool["Machinery Location"])


def F(df):
    """Apply sidebar filters to any result table."""
    if df is None or df.empty:
        return df
    out = df
    if ss.f_func:
        if "Function" in out.columns:
            out = out[out["Function"].isin(ss.f_func)]
        elif "Machinery Location" in out.columns:
            out = out[out["Machinery Location"].isin(mach_in_func)]
    if ss.f_mach and "Machinery Location" in out.columns:
        out = out[out["Machinery Location"].isin(ss.f_mach)]
    if ss.f_search:
        cols = [c for c in ["Title", "Sub Component Location", "Job Code", "Machinery Location"] if c in out.columns]
        if cols:
            mask = pd.Series(False, index=out.index)
            for c in cols:
                mask |= out[c].astype(str).str.contains(ss.f_search, case=False, regex=False)
            out = out[mask]
    return out


def header(title, hint):
    st.header(title)
    st.markdown(f"<div class='hint'>{hint}</div>", unsafe_allow_html=True)
    chips = []
    if ss.f_func:
        chips.append("Function: " + ", ".join(ss.f_func))
    if ss.f_mach:
        chips.append("Machinery: " + ", ".join(ss.f_mach))
    if ss.f_search:
        chips.append(f"Search: “{ss.f_search}”")
    if chips:
        c1, c2 = st.columns([6, 1])
        c1.info("Filtered by  " + "  ·  ".join(chips), icon="🔎")
        c2.button("Clear filters", on_click=clear_filters, key=f"clr_{title}")


def table(df, name, height=430, **kw):
    if df is None or df.empty:
        st.success("No differences found ✅")
        return
    st.caption(f"{len(df):,} rows · click a column header to sort · hover the table for search / full-screen")
    st.dataframe(df, use_container_width=True, hide_index=True, height=height, **kw)
    st.download_button("⬇ Download this table (CSV)", df.to_csv(index=False).encode("utf-8-sig"),
                       file_name=f"{name.replace(' ', '_')}.csv", mime="text/csv", key=f"csv_{name}")


def hbar(df, color, height=340):
    by = df.groupby("Machinery Location").size().sort_values(ascending=False).head(12).reset_index(name="Jobs")
    fig = px.bar(by, x="Jobs", y="Machinery Location", orientation="h", height=height, color_discrete_sequence=[color])
    fig.update_layout(yaxis=dict(autorange="reversed", title=None), margin=dict(l=0, r=0, t=10, b=0))
    st.plotly_chart(fig, use_container_width=True)


page = ss.nav

# ====================================================================
# Summary
# ====================================================================
if page == "summary":
    header(f"{NA}  vs  {NB}", "Start here. Each card is one type of discrepancy – click <b>Open</b> to see the details.")

    S = R["summary"].set_index("Metric")
    rows = [("Total jobs", "Total jobs"), ("Machinery", "Machinery Locations"),
            ("Sub-components", "Sub Component Locations (machinery+sub)")]
    cols = st.columns(len(rows))
    for col, (label, key) in zip(cols, rows):
        va, vb = int(S.loc[key, NA]), int(S.loc[key, NB])
        with col, st.container(border=True):
            st.markdown(f"**{label}**")
            x, y = st.columns(2)
            x.metric(f"A", f"{va:,}")
            y.metric(f"B", f"{vb:,}", delta=vb - va if vb != va else None)

    st.subheader("Key findings")
    cards = [
        ("🔢", "Machinery with different job count", int((R["Machinery Count"]["Status"] != "Match").sum()),
         "Same machinery, but one vessel has more jobs on it.", "counts", {"cnt_level": "Machinery Location"}),
        ("❓", f"Jobs only on A ({NA})", len(R["Jobs Only A"]),
         f"{int(R['Jobs Only A']['Likely cause'].str.startswith('Same').sum())} of these exist on B under a different sub-component.",
         "missing", {"miss_view": "A"}),
        ("❓", f"Jobs only on B ({NB})", len(R["Jobs Only B"]),
         f"{int(R['Jobs Only B']['Likely cause'].str.startswith('Same').sum())} of these exist on A under a different sub-component.",
         "missing", {"miss_view": "B"}),
        ("⏱", "Different frequency", len(R["Frequency Diff"]),
         "Same job, different interval (months / running hours).", "frequency", {}),
        ("👷", "Different performing rank", len(R["PerfRank Diff"]),
         "Same job assigned to a different rank (order ignored).", "rank", {}),
        ("📝", "Different description", len(R["Description Diff"]),
         "Job instructions worded differently – see side-by-side view.", "description", {}),
        ("🚩", "Different Critical flag", int(other_counts.get("Critical", 0)),
         "Job marked critical on one vessel only.", "other", {"other_field": "Critical"}),
        ("⚙", "Different Maker / Model", len(R["Maker Model Diff"]),
         "Machinery recorded with different maker or model.", "equipment", {}),
        ("🧹", "Sub-component naming issue", len(R["Naming Issue"]),
         "Names repeating the machinery prefix – should be fixed in PMS.", "equipment", {}),
    ]
    for i in range(0, len(cards), 3):
        cols = st.columns(3)
        for col, (icon, title, n, desc, target, preset) in zip(cols, cards[i:i + 3]):
            with col, st.container(border=True):
                st.markdown(f"#### {icon} {n:,}")
                st.markdown(f"**{title}**")
                st.caption(desc)
                st.button("Open →", key=f"card_{title}", on_click=go, args=(target,), kwargs=preset,
                          disabled=n == 0, use_container_width=True)

    st.subheader("Where are the most discrepancies?")
    st.markdown("<div class='hint'>Select a machinery in the table to open its drill-down.</div>", unsafe_allow_html=True)
    hot = F(R["Hotspots"])
    hot = hot[hot["Total discrepancies"] > 0]
    l, r = st.columns([1, 1])
    with l:
        top = hot.head(15)
        if len(top):
            vcols = [c for c in top.columns if c not in ("Machinery Location", "Function", "Total discrepancies")]
            long = top.melt(id_vars="Machinery Location", value_vars=vcols, var_name="Type", value_name="Count")
            fig = px.bar(long[long["Count"] > 0], y="Machinery Location", x="Count", color="Type", orientation="h",
                         height=520, category_orders={"Machinery Location": top["Machinery Location"].tolist()})
            fig.update_layout(legend=dict(orientation="h", y=-0.2, title=None), yaxis_title=None,
                              margin=dict(l=0, r=0, t=10, b=0))
            st.plotly_chart(fig, use_container_width=True)
    with r:
        ev = st.dataframe(hot[["Machinery Location", "Function", "Total discrepancies"]], use_container_width=True,
                          hide_index=True, height=470, on_select="rerun", selection_mode="single-row", key="hot_tbl")
        sel = ev.selection.rows if ev and ev.selection else []
        if sel:
            m = hot.iloc[sel[0]]["Machinery Location"]
            st.button(f"🔍 Open drill-down: {m}", type="primary", on_click=go, args=("drill",),
                      kwargs={"drill_mach": m}, use_container_width=True)

# ====================================================================
# Counts
# ====================================================================
elif page == "counts":
    header("🔢 Count differences", "Number of jobs per level on each vessel. "
           "<b>Count differs</b> = both vessels have it but with a different number of jobs; "
           "<b>Only in …</b> = exists on one vessel only.")
    ss.setdefault("cnt_level", "Machinery Location")
    lvl = st.segmented_control("Level", ["Function", "Machinery Location", "Sub-component"], key="cnt_level") or "Machinery Location"
    key = {"Function": "Function Count", "Machinery Location": "Machinery Count", "Sub-component": "SubComponent Count"}[lvl]
    df = F(R[key])
    statuses = ["Count differs", f"Only in {NA}", f"Only in {NB}", "Match"]
    sc = df["Status"].value_counts()
    cc = st.columns(4)
    for i, s in enumerate(statuses):
        cc[i].metric(s.replace(NA, "A").replace(NB, "B"), int(sc.get(s, 0)))
    show_match = st.toggle("Also show rows that match", value=False)
    if not show_match:
        df = df[df["Status"] != "Match"]
    table(df, f"{lvl} counts")

# ====================================================================
# Missing jobs
# ====================================================================
elif page == "missing":
    header("❓ Missing jobs", "Jobs are matched on <b>Machinery + Sub-component + Job Code</b>. "
           "<b>Likely cause</b> tells you whether the job really is missing, or exists on the same machinery "
           "under a different sub-component (usually a numbering / naming difference).")
    ss.setdefault("miss_view", "A")
    view = st.segmented_control("Show", ["A", "B", "codes"], key="miss_view",
                                format_func=lambda v: {"A": f"Only on A – {NA} ({len(R['Jobs Only A'])})",
                                                       "B": f"Only on B – {NB} ({len(R['Jobs Only B'])})",
                                                       "codes": f"Job codes not used on other vessel ({len(R['JobCode Missing'])})"}[v]) or "A"
    if view == "codes":
        table(R["JobCode Missing"], "JobCode_Missing")
    else:
        df = F(R["Jobs Only A" if view == "A" else "Jobs Only B"])
        if df is not None and len(df):
            cause = df["Likely cause"].value_counts()
            c = st.columns(len(cause))
            for col, (k, v) in zip(c, cause.items()):
                col.metric(k, v)
            only_real = st.toggle("Show only jobs genuinely missing on the machinery", value=False)
            if only_real:
                df = df[df["Likely cause"] == "Job missing on this machinery"]
            l, r = st.columns([2, 1])
            with l:
                table(df, f"Jobs_only_{view}")
            with r:
                st.markdown("**Top machinery**")
                if len(df):
                    hbar(df, COLOR_A if view == "A" else COLOR_B)
        else:
            table(df, f"Jobs_only_{view}")

# ====================================================================
# Frequency
# ====================================================================
elif page == "frequency":
    header("⏱ Frequency differences", "Same job on both vessels but with a different interval. "
           "The <b>patterns</b> table shows changes that repeat across many jobs – often a single root cause.")
    df = F(R["Frequency Diff"])
    if df is None or df.empty:
        table(df, "Frequency_Diff")
    else:
        l, r = st.columns([1, 1])
        with l:
            st.markdown("**Recurring patterns**")
            pat = (df.groupby(["Frequency" + sa, "Frequency" + sb]).size().reset_index(name="Jobs")
                   .sort_values("Jobs", ascending=False))
            pat.columns = ["Frequency on A", "Frequency on B", "Jobs"]
            st.dataframe(pat, use_container_width=True, hide_index=True, height=340)
        with r:
            st.markdown("**Top machinery**")
            hbar(df, COLOR_A)
        st.markdown("**All jobs with different frequency**")
        table(df, "Frequency_Diff")

# ====================================================================
# Performing rank
# ====================================================================
elif page == "rank":
    header("👷 Performing rank differences", "Same job assigned to a different performing rank. "
           "The order in which ranks are listed is ignored.")
    df = F(R["PerfRank Diff"])
    if df is not None and len(df):
        pat = (df.groupby(["Performing Rank" + sa, "Performing Rank" + sb]).size().reset_index(name="Jobs")
               .sort_values("Jobs", ascending=False))
        pat.columns = ["Rank on A", "Rank on B", "Jobs"]
        st.markdown("**Recurring patterns**")
        st.dataframe(pat, use_container_width=True, hide_index=True)
    table(df, "PerfRank_Diff")

# ====================================================================
# Description
# ====================================================================
elif page == "description":
    header("📝 Description differences", "Each job gets a <b>similarity %</b>: 0 % = completely different text, "
           "close to 100 % = small wording change. <b>Select a row</b> to compare both descriptions side-by-side.")
    df = F(R["Description Diff"])
    if df is None or df.empty:
        table(df, "Description_Diff")
    else:
        band = st.segmented_control("Show", ["All", "Very different (< 30 %)", "Partly different (30–80 %)",
                                             "Minor wording (> 80 %)"], default="All", key="desc_band") or "All"
        rng = {"All": (0, 101), "Very different (< 30 %)": (0, 30), "Partly different (30–80 %)": (30, 80),
               "Minor wording (> 80 %)": (80, 101)}[band]
        df = df[(df["Similarity %"] >= rng[0]) & (df["Similarity %"] < rng[1])].reset_index(drop=True)
        view_cols = ["Machinery Location", "Sub Component Location", "Job Code", "Title", "Similarity %"]
        st.caption(f"{len(df):,} jobs")
        ev = st.dataframe(df[view_cols], use_container_width=True, hide_index=True, height=330,
                          on_select="rerun", selection_mode="single-row", key="desc_tbl",
                          column_config={"Similarity %": st.column_config.ProgressColumn(
                              "Similarity %", min_value=0, max_value=100, format="%.0f%%")})
        rows = ev.selection.rows if ev and ev.selection else []
        if rows:
            rec = df.iloc[rows[0]]
            with st.container(border=True):
                st.markdown(f"#### {rec['Title']}  ·  Job {rec['Job Code']}  ·  {rec['Similarity %']:.0f}% similar")
                st.caption(f"{rec['Machinery Location']}  ›  {rec['Sub Component Location']}")
                st.markdown("<span style='background:#ffd6d6;padding:1px 6px'>red</span> = text only on A &nbsp; "
                            "<span style='background:#d4f5d4;padding:1px 6px'>green</span> = text only on B",
                            unsafe_allow_html=True)
                ha, hb = diff_html(rec["Description" + sa], rec["Description" + sb])
                ca, cb = st.columns(2)
                box = ("<div style='border:1px solid #ccc;border-radius:6px;padding:10px;max-height:480px;"
                       "overflow:auto;font-size:0.9rem;line-height:1.5'>{}</div>")
                ca.markdown(f"<span class='badge' style='background:{COLOR_A}'>A</span> **{NA}**", unsafe_allow_html=True)
                ca.markdown(box.format(ha or "<i>(empty)</i>"), unsafe_allow_html=True)
                cb.markdown(f"<span class='badge' style='background:{COLOR_B}'>B</span> **{NB}**", unsafe_allow_html=True)
                cb.markdown(box.format(hb or "<i>(empty)</i>"), unsafe_allow_html=True)
        else:
            st.info("👆 Select a row (tick box on the left of the table) to see the side-by-side comparison.")
        st.download_button("⬇ Download descriptions (CSV, full text)", df.to_csv(index=False).encode("utf-8-sig"),
                           file_name="Description_Diff.csv", mime="text/csv")

# ====================================================================
# Other fields
# ====================================================================
elif page == "other":
    header("🏷 Critical flag & other fields", "Differences in Critical flag, Verifying Rank, Title, Job Source and E-Form "
           "for jobs that exist on both vessels.")
    df = F(R["Other Field Diff"])
    fields = ["Critical", "Verifying Rank", "Title", "Job Source", "E-Form"]
    cnt = df["Field"].value_counts() if df is not None and len(df) else pd.Series(dtype=int)
    ss.setdefault("other_field", "Critical")
    fsel = st.segmented_control("Field", fields, key="other_field",
                                format_func=lambda f: f"{f} ({int(cnt.get(f, 0))})") or "Critical"
    sub = df[df["Field"] == fsel] if df is not None and len(df) else df
    if sub is not None and len(sub):
        summ = sub[[f"Value{sa}", f"Value{sb}"]].replace("", "(blank)").value_counts().rename("Jobs").reset_index()
        summ.columns = ["Value on A", "Value on B", "Jobs"]
        st.dataframe(summ, hide_index=True)
    table(sub, f"Other_{fsel}")

# ====================================================================
# Equipment & naming
# ====================================================================
elif page == "equipment":
    header("⚙ Maker/Model & naming", "Equipment and data-quality differences that can explain why job lists differ.")
    st.subheader(f"Maker / Model differences ({len(F(R['Maker Model Diff']))})")
    table(F(R["Maker Model Diff"]), "Maker_Model_Diff", height=280)
    st.subheader(f"Sub-component naming issue ({len(F(R['Naming Issue']))})")
    st.markdown("<div class='hint'>Sub-component names that repeat the machinery as a prefix, e.g. "
                "<code>Auxiliary Boiler#1 &gt; Auxiliary Boiler - General#1</code>. The tool strips the prefix before "
                "matching; these names should be corrected in the PMS.</div>", unsafe_allow_html=True)
    table(F(R["Naming Issue"]), "Naming_Issue", height=330)

# ====================================================================
# Machinery drill-down
# ====================================================================
elif page == "drill":
    header("🔍 Machinery drill-down", "Every discrepancy for one machinery on a single page. "
           "The list is sorted by number of discrepancies.")
    hot = F(R["Hotspots"])
    options = hot["Machinery Location"].tolist() if hot is not None and len(hot) else []
    tot = dict(zip(R["Hotspots"]["Machinery Location"], R["Hotspots"]["Total discrepancies"]))
    if not options:
        st.warning("No machinery for the current filter.")
        st.stop()
    if ss.get("drill_mach") not in options:
        ss.drill_mach = options[0]
    mach = st.selectbox("Machinery Location", options, key="drill_mach",
                        format_func=lambda m: f"{m}   —   {tot.get(m, 0)} discrepancies")
    mc = R["Machinery Count"]
    x = mc[mc["Machinery Location"] == mach].iloc[0]
    k = st.columns(4)
    k[0].metric("Jobs on A", int(x[NA]))
    k[1].metric("Jobs on B", int(x[NB]), delta=int(x[NB] - x[NA]) or None)
    k[2].metric("Discrepancies", tot.get(mach, 0))
    k[3].metric("Function", x["Function"])

    sections = [("Sub-component count differences", "SubComponent Count"), (f"Jobs only on A", "Jobs Only A"),
                (f"Jobs only on B", "Jobs Only B"), ("Frequency differences", "Frequency Diff"),
                ("Performing rank differences", "PerfRank Diff"), ("Description differences", "Description Diff"),
                ("Critical & other field differences", "Other Field Diff"), ("Maker / Model", "Maker Model Diff")]
    for title, key in sections:
        d = R[key]
        d = d[d["Machinery Location"] == mach]
        if key == "SubComponent Count":
            d = d[d["Status"] != "Match"]
        icon = "🟠" if len(d) else "🟢"
        with st.expander(f"{icon} {title} ({len(d)})", expanded=0 < len(d) <= 25):
            if len(d):
                st.dataframe(d, use_container_width=True, hide_index=True)
            else:
                st.write("No differences ✅")

# ====================================================================
# Export
# ====================================================================
elif page == "export":
    header("⬇ Export report", "Download the comparison as a formatted Excel workbook "
           "(Summary sheet + one sheet per check, with filters and frozen headers).")
    with st.container(border=True):
        st.markdown("**1. Choose sheets**")
        exp_sheets = st.multiselect("Sheets", SHEET_ORDER, default=SHEET_ORDER, label_visibility="collapsed")
        st.markdown("**2. Options**")
        apply_filters = st.checkbox("Apply the current filters to the export", value=False,
                                    disabled=not (ss.f_func or ss.f_mach or ss.f_search))
        st.markdown("**3. Build and download**")
        if st.button("📊 Build Excel report", type="primary", disabled=not exp_sheets):
            with st.spinner("Building report…"):
                data = {k: (F(v) if apply_filters and isinstance(v, pd.DataFrame) and k in SHEET_ORDER else v)
                        for k, v in R.items()}
                ss.xlsx = build_excel(data, exp_sheets)
        if "xlsx" in ss:
            st.download_button("⬇ Download Excel report", ss.xlsx, type="primary",
                               file_name=f"JD_Comparison_{NA}_vs_{NB}.xlsx".replace(" ", "_"),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("Tip: every table in the tool also has its own CSV download button.")
