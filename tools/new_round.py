"""Start a new LMD round from its distribution plan.

Usage (from the project folder):
    py tools/new_round.py "private/plan.xlsx" "LMD Sept-Oct 2026"

Reads the 'FP LMD Order' sheet and writes data/plan.csv: every facility with
something planned this round (names only, no quantities).
"""
import csv
import json
import os
import sys

import openpyxl

from common import DATA, find_facility, load_facilities, load_names

# Short names for the commodities, matched on how the plan's name starts
SHORT = {"condom female": "Female condom", "condom male": "Male condom", "depo-provera": "Depo-Provera",
         "exluton": "Exluton", "iucd": "IUCD", "lo-femenal": "Lo-Femenal", "microgynon": "Microgynon",
         "noristerat": "Noristerat", "implanon": "Implanon NXT", "jadelle": "Jadelle",
         "cyclebeads": "Cycle Beads", "consumables kit": "Consumables kit", "sayana": "Sayana Press",
         "levoplant": "Levoplant", "disposable trocar": "Disposable Trocar"}


def short_name(name):
    lower = str(name).strip().lower()
    for start, short in SHORT.items():
        if lower.startswith(start):
            return short
    return str(name).strip()


def read_plan(path):
    sheet = openpyxl.load_workbook(path, data_only=True, read_only=True)["FP LMD Order"]
    rows = list(sheet.iter_rows(values_only=True))

    # Row 5 has the column headings; the commodity names are 2 rows above them
    headings = [str(h or "").strip() for h in rows[4]]
    commodity_row = rows[2]
    site_col, lga_col = headings.index("SiteName"), headings.index("LGA")

    # Each commodity has 7 columns; the planned amount is its 'Validated Order Quantity'
    commodities = []
    for col, heading in enumerate(headings):
        if heading == "Validated Order Quantity":
            name_col = max(c for c in range(col + 1) if commodity_row[c])
            commodities.append((col, short_name(commodity_row[name_col])))

    plan = []
    for row in rows[5:]:
        if not row[site_col] or not row[lga_col]:
            continue
        amounts = {name: int(row[col]) for col, name in commodities if row[col]}
        if amounts:  # skip facilities with nothing planned this round
            plan.append((row[lga_col], row[site_col], amounts))
    return plan, [name for _, name in commodities]


def main():
    plan_path, round_name = sys.argv[1], sys.argv[2]
    facilities, names = load_facilities(), load_names()
    plan, commodities = read_plan(plan_path)

    matched, problems = [], []
    for lga, site, amounts in plan:
        facility = find_facility(site, lga, facilities, names)
        if facility:
            matched.append((facility, amounts))
        else:
            problems.append(f"{lga} | {site}")

    if problems:
        print("These facilities in the plan are not in data/facilities.csv.")
        print("Add each one to data/names.csv (LGA, Name as written, S/N), then run this again:")
        for p in problems:
            print("   ", p)
        sys.exit(1)

    with open(os.path.join(DATA, "plan.csv"), "w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["S/N", "LGA", "Facility Name"])  # no quantities: this file goes on GitHub
        for facility, amounts in matched:
            writer.writerow([facility["sn"], facility["lga"], facility["name"]])

    with open(os.path.join(DATA, "round.json"), "w", encoding="utf-8") as fh:
        json.dump({"round": round_name}, fh, indent=1)

    print(f"{round_name}: {len(matched)} facilities have something planned. Saved data/plan.csv")


if __name__ == "__main__":
    main()