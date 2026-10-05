# Sister Vessel JD Comparison

Compares two PMS **Job Status (with Job Details)** extracts from sister vessels and finds the discrepancies.

## Run the tool
Double-click `run_app.bat`, or:

```
python -m streamlit run app.py --server.port 8531
```

**Step 1:** upload the two extracts (or paste their file paths). **Step 2:** click **Compare vessels**.

## Navigation
The sidebar lists every page with its number of discrepancies. Start on **Summary**: each finding card has an
**Open →** button that jumps to the detail page, and selecting a machinery in the hotspot table opens its drill-down.

| Page | Shows |
|---|---|
| Summary | A vs B totals, finding cards, discrepancy hotspots |
| Count differences | Job counts per Function / Machinery / Sub-component |
| Missing jobs | Jobs on only one vessel (with likely cause), job codes not used on the other vessel |
| Frequency | Jobs with different frequency + recurring patterns (e.g. 18000 h → 36000 h) |
| Performing rank | Rank differences (order-insensitive) |
| Description | Similarity bands; select a row for a side-by-side highlighted comparison |
| Critical & other fields | Critical, Verifying Rank, Title, Job Source, E-Form differences |
| Maker/Model & naming | Maker/Model differences; sub-component names with a duplicated machinery prefix |
| Machinery drill-down | Every discrepancy for one machinery |
| Export report | Choose sheets and download the Excel report |

**Filters** (sidebar: Function, Machinery, text search) apply to every page; active filters are shown at the top
of the page with a **Clear filters** button. Every table has its own CSV download.

## Matching logic
- Jobs are matched on **Machinery Location + Sub Component Location + Job Code**; duplicates pair up 1:1.
- Text is normalised (whitespace, `&amp;`) before comparing; descriptions are compared case-insensitively.
- Sub-components written as `<Machinery> > <Sub-component>` are normalised to `<Sub-component>` before matching
  and listed on the *Naming Issue* sheet.

## Command line (no UI)
```
python compare_cli.py "<vesselA.xlsx>" "<vesselB.xlsx>" [output.xlsx]
```

## Files
- `compare_engine.py` – comparison logic and Excel builder (shared by the app and CLI)
- `app.py` – Streamlit UI
- `compare_cli.py` – command-line report
