# Job Status + Job List Merger

Adds **Job Description, Maker and Model** from a PMS *Job List* export to every row of a
PMS *Job Status* export.

## Run the app
Double-click `run_app.bat`, or:

```
pip install -r requirements.txt
python -m streamlit run app.py
```

Upload the Job Status file and the Job List file (CSV or Excel), review the match summary,
and download the merged Excel/CSV.

## Command line (no app)
```
python merge_logic.py "<Job Status file>" "<Job List file>" [output.xlsx|output.csv]
```

## How rows are matched
One Job Code can be linked to several machineries, so the match uses
**Job Code + Machinery + Component**. Before comparing, names are cleaned:
upper/lower case and spaces ignored, `Starboard/Stbd/Port/Forward/Fwd/Aft` → `S/P/F/A`,
the `#1` ending dropped from components, only the last part of `Parent > Child` used,
and `High Velocity PV Valve-Cargo Tank S3` treated as matching `High Velocity PV Valve`.

| Match Level | Meaning |
|---|---|
| 1 – Code + Machinery + Component | Most reliable |
| 2 – Code + Machinery | One Job List entry for that code/machinery (or all entries agree) |
| 3 – ambiguous – review | Several Job List entries with different maker/model; all values shown, separated by ` \| ` |
| 4 – Code only | Machinery name differs; code exists on one machinery only |
| Not Found | No matching Job Code/Machinery |

The **Match Note** column explains level 3, 4 and Not Found rows.
