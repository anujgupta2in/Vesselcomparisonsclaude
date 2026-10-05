"""Command-line report: python compare_cli.py <vesselA.xlsx> <vesselB.xlsx> [output.xlsx]"""
import sys

from compare_engine import read_extract, compare, build_excel

if len(sys.argv) < 3:
    sys.exit(__doc__)

res = compare(read_extract(sys.argv[1]), read_extract(sys.argv[2]))
na, nb = res["names"]
out = sys.argv[3] if len(sys.argv) > 3 else f"JD_Comparison_{na}_vs_{nb}.xlsx".replace(" ", "_")
with open(out, "wb") as f:
    f.write(build_excel(res))
print(res["summary"].to_string(index=False), "\n")
print(res["findings"][["Check", "Count"]].to_string(index=False))
print(f"\nReport written to {out}")
