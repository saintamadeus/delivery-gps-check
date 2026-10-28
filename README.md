# LMD Delivery Status

Shows which health facilities in Lagos have received their family planning commodities in the current
last-mile distribution (LMD) round, and which still need a follow-up call.

**Live site:** https://saintamadeus.github.io/delivery-gps-check/

## How it works

1. The delivery person fills a short **KoBo form** at the facility: their name, the LGA, the facility (picked
   from a list), and the facility's **6-digit code**.
2. The facility in-charge gives the code **only after the goods have arrived**. The form checks the code on
   the spot and records the phone's location by itself (it cannot be typed or moved).
3. If the in-charge is not there, the delivery person ticks "In-charge not available" and writes who received
   the goods instead. If they are more than 500 m from the facility, they must say why.
4. **Every hour**, GitHub fetches the forms from KoBo, works out each facility's status, and updates the website
   and the Excel download.

## What the statuses mean

| Status | Meaning |
| --- | --- |
| ✅ Receipt confirmed by facility | The right facility code was entered. ⚠️ means it was entered far from the facility. |
| 🟡 Reported delivered, location matches | No code (in-charge not available), but the phone was within 500 m of the facility. |
| 🔴 Reported delivered, location doesn't match | No code, and the phone was more than 500 m away, or no location was recorded. |
| ⚪ Not reported yet | No form has been filled for this facility. |

The website shows **no names, phone numbers, quantities or codes**. Reasons typed by delivery people are only
visible in KoBo.

## Starting a new round

Run these from the project folder on the laptop (Windows PowerShell).

1. Save the new distribution plan as `private\plan.xlsx`.
2. Read the plan:
   ```
   py tools/new_round.py "private/plan.xlsx" "LMD Nov-Dec 2026"
   ```
   If it lists facilities it cannot find, add each one to `data/names.csv` (LGA, name as written in the plan,
   S/N from `data/facilities.csv`) and run it again.
3. Make new codes: `py tools/make_codes.py`
   This writes one Excel file per LGA in `private\<round>\`. Send each file **privately** to that LGA's FP
   coordinator.
4. Make the form: `py tools/make_form.py`
   In KoBo, open the form, click **⋯ → Replace form**, pick `private\<round>\delivery_form.xlsx`, then
   **Redeploy**. The form link stays the same.
5. Save and upload:
   ```
   git pull
   git add -A
   git commit -m "New round: LMD Nov-Dec 2026"
   git push
   ```

## Files

| File | What it is |
| --- | --- |
| `index.html` | The website |
| `data/facilities.csv` | Every facility: S/N, LGA, name, latitude, longitude |
| `data/names.csv` | Other spellings of facility names used in plans, linked to the right S/N |
| `data/plan.csv` | Facilities expecting goods this round (names only, no quantities) |
| `data/round.json` | The current round's name |
| `data/status.csv`, `data/status.xlsx` | Each facility's status (updated every hour) |
| `data/updated.json` | When the status was last updated |
| `tools/` | The Python scripts above, plus `update.py` (the hourly status check) |
| `.github/workflows/update.yml` | Tells GitHub to run `update.py` every hour |

**Never uploaded** (the `private` folder is ignored by git): the distribution plan, the codes, the KoBo form
file, `kobo.json` (the KoBo key), and the full form answers with names.

## Settings on GitHub

The hourly update needs two secrets (**Settings → Secrets and variables → Actions**):
`KOBO_FORM` (the form's ID) and `KOBO_TOKEN` (the KoBo API key). To run an update straight away:
**Actions → Hourly update → Run workflow**.

## Sources

Facility names and LGAs: the Lagos Family Planning Last Mile Distribution Plan. Coordinates: GRID3 NGA Health
Facilities v2.0 (Nigeria Health Facility Registry, CC BY 4.0), Google Maps and OpenStreetMap, checked by hand.