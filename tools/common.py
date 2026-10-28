"""Small helpers shared by all the tools."""
import csv
import os

# The project folder, worked out from where this file sits (tools/ is inside it)
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data")


def lga_key(text):
    """Make LGA names comparable: 'Amuwo_Odofin', 'amuwo odofin', 'Amuwo-Odofin' all become 'amuwo odofin'."""
    text = str(text or "").lower()
    for ch in "_-/":
        text = text.replace(ch, " ")
    return " ".join(text.split())


def name_key(text):
    """Make facility names comparable: ignore capitals and extra spaces."""
    return " ".join(str(text or "").lower().split())


def load_facilities():
    """Read data/facilities.csv into a list of dicts: sn, lga, name, lat, lon."""
    facilities = []
    with open(os.path.join(DATA, "facilities.csv"), encoding="utf-8-sig", newline="") as fh:
        for row in csv.DictReader(fh):
            def number(v):
                try:
                    return float(v)
                except (TypeError, ValueError):
                    return None  # blank = no location yet
            facilities.append({
                "sn": row["S/N"].strip(),
                "lga": row["LGA"].strip(),
                "name": row["Facility Name"].strip(),
                "lat": number(row["Latitude"]),
                "lon": number(row["Longitude"]),
            })
    return facilities


def load_names():
    """Read data/names.csv: other spellings of a facility, linked to its S/N."""
    names = {}
    with open(os.path.join(DATA, "names.csv"), encoding="utf-8-sig", newline="") as fh:
        for row in csv.DictReader(fh):
            names[(lga_key(row["LGA"]), name_key(row["Name as written"]))] = row["S/N"].strip()
    return names


def find_facility(name, lga, facilities, names):
    """Find a facility by its name inside its LGA. Returns the facility, or None if not found.

    The same name can exist in two LGAs (e.g. PHC Ijegun), so we always look inside the LGA.
    """
    key = (lga_key(lga), name_key(name))
    if key in names:  # a known other spelling
        sn = names[key]
        return next((f for f in facilities if f["sn"] == sn), None)
    for f in facilities:
        if lga_key(f["lga"]) == key[0] and name_key(f["name"]) == key[1]:
            return f
    return None