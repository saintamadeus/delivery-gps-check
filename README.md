# Delivery GPS Check

Checks each Lagos family planning last-mile delivery against the real location of its health facility, and flags deliveries recorded too far away, with no GPS, or from a GPS spot shared by several facilities.

Live site: https://saintamadeus.github.io/delivery-gps-check/

## Using it

1. Export the delivery form data as Excel (.xlsx) or CSV.
2. Open the live site and click **Upload tracker file**.
3. Read the result tiles, the map and the list. Click **Download results (Excel)** to keep a copy.

The tracker file is read inside your browser. It is never uploaded to GitHub or anywhere else.

## Files

| File | What it is |
| --- | --- |
| `index.html` | The tool |
| `facilities.csv` | The shared facility list: S/N, LGA, name, latitude, longitude, status, source |
| `aliases.csv` | Facility names as typed on the phone, linked to the right S/N |
| `lib/` | Leaflet 1.9.4 (map) and SheetJS 0.18.5 (Excel), kept here so the site does not depend on another server |

## Updating facility locations

Either:

- edit `facilities.csv` directly on GitHub (click the file, then the pencil icon), and commit; or
- make the changes in the tool (Facility locations tab), click **Download for GitHub**, then upload the downloaded `facilities.csv` and `aliases.csv` here (**Add file → Upload files**), replacing the old ones, and commit.

The live site picks up a commit within a few minutes.

Coordinates are decimal degrees, latitude first (Lagos is about 6.4 to 6.7, and 2.7 to 4.3).

## Sources

Facility names and LGAs: Sept–Oct 2026 Family Planning Last Mile Distribution Plan. Coordinates: GRID3 NGA Health Facilities v2.0 (Nigeria Health Facility Registry, CC BY 4.0), Google Maps and OpenStreetMap, checked by hand.
