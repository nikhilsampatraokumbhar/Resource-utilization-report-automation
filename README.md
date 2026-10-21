# Resource utilisation report automation

Automates the PMO's daily resource-utilisation report:

1. Every weekday at 9:00, a Power Automate flow fetches the previous working day's logged hours from Clockify.
2. It groups people under their Team Leads using the TL–Resource mapping in Excel.
3. It calculates each person's logged hours, capacity and utilisation %, taking leave and public holidays into account.
4. It posts a separate report in each Team Lead's Microsoft Teams channel.

After setup, day-to-day maintenance happens in the Excel file: adding people, moving them between TLs, recording leave.

## What's in this repo

| Path | What it is |
| --- | --- |
| `docs/power-automate-guide.html` | Step-by-step build guide with illustrations and copy-ready formulas. Open it in a browser. |
| `templates/Utilisation-Mapping-Template.xlsx` | The Excel workbook the flow reads (tables `Resources`, `TeamLeads`, `Leave`, `History`). Upload it to OneDrive for Business or SharePoint and replace the example rows. |
| `scripts/build_template.py` | Regenerates the Excel template (`uv run --with openpyxl python scripts/build_template.py`). |

## Requirements

- Power Automate with a Premium licence (or the 90-day Premium trial) for the flow owner only. The HTTP action that calls Clockify is a Premium connector; everything else uses standard Microsoft 365 connectors.
- A Clockify API key from an Admin or Owner account, so the report covers everyone.
- The HTTP connector allowed by the company's Power Automate data policy. Check with IT.

## How utilisation is calculated

```
Utilisation % = Logged hours ÷ Capacity × 100
Capacity      = Daily Hours from Excel (blank = 8), or 0 when the person is on leave
```

| Status | Rule (default) |
| --- | --- |
| ✅ On track | 80–110% |
| 🔻 Under-utilised | below 80% |
| 🔺 Over-utilised | above 110% |
| ⚠️ No time logged | 0 h logged |
| 🌴 On leave | listed in the Leave table |
| ❓ Not in Clockify | email in Excel doesn't match a Clockify user |
