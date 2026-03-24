# ATTAIR Fleet Reliability Report Generator

**CAME Part 5 — Reliability Control Program, Issue 2 Rev 0, 30/08/2024**

Generates a fully formatted `.docx` fleet-level reliability report covering multiple aircraft in a single document. All 13 CAME-required sections are produced, including charts (via QuickChart.io), Alert Notice boxes, and a post-generation compliance checklist.

---

## Requirements

- Node.js ≥ 18
- npm

## Setup

```bash
npm install
```

---

## Quick Start

### 1. Configure the report

Edit **`fleet-config.js`** and fill in:

| Field | Description |
|---|---|
| `period.start` / `period.end` | Reporting period dates |
| `period.monthsInScope` | Array of months, e.g. `["January 2026", "February 2026", "March 2026"]` |
| `period.label` | Human-readable label printed on the cover |
| `preparedBy` | Your name and title |
| `aircraft[]` | One entry per aircraft (see below) |

**Per-aircraft fields:**

```js
{
  registration: "5N-BZN",
  msn:           "145556",
  type:          "Embraer ERJ-145",
  variant:       "EMB-145MP",
  engines:       "auto-detect",     // or "2 × Rolls-Royce AE3007A1P"
  apu:           "auto-detect",     // "fitted" | "not fitted" | "auto-detect"
  note:          "",                // e.g. "In heavy check Feb 2026"
  files: {
    defectRecords:     "./data/5N-BZN/defect-records.xlsx",
    deferredDefects:   "./data/5N-BZN/mel-sheet.xlsx",
    tlp:               "./data/5N-BZN/tlp.xlsx",
    fuelOil:           "./data/5N-BZN/fuel-oil-log.xlsx",
    delayCancellation: "./data/5N-BZN/delay-cancel.xlsx",  // optional — set null if absent
    pirepMarep:        "./data/5N-BZN/pirep-marep.xlsx",   // optional — set null if absent
  },
  dataGaps: [
    // { item: "Fuel & Oil", month: "January 2026", reason: "Not recorded" }
  ],
}
```

> Set any `files` entry to `null` if the file is not available. Missing files produce `"N/A"` / `"Not recorded"` entries in the report — they are **not** treated as compliance failures.

### 2. Organise your data files

Place files in the paths you specified. Supported formats: **`.xlsx`**, **`.xls`** (Excel workbooks).

**Expected file types:**

| File | Content |
|---|---|
| TLP | Technical Log Page — one sheet per month. Rows contain `FROM`/`TO` columns for each sector. `B/FWD` row = opening values; `TOTAL` row = closing values. |
| Defect Records | Defect register with `ATA`, `Description`, `Date`, `Status` columns. |
| MEL / Deferred Defects | MEL register with `ATA`, `Category`, `Raised Date`, `Due Date`, `Status`. |
| Fuel & Oil Log | Monthly summary with `Month`, `FH`, fuel uplift/burn, engine oil uplift/consumption/rate columns. First 5 rows may contain engine SNs and oil limits (auto-detected). |
| Delay & Cancellation | Log with `Date`, `Type` (DELAY/CANCELLATION), `ATA`, delay minutes. |
| PIREP / MAREP Register | Log with `Date`, `Type` (PIREP/MAREP), `ATA`, description. |

> The parsers use flexible column header matching (case-insensitive, partial matches) so your exact column names do not need to match these exactly.

### 3. Generate the report

```bash
node generate-fleet-report.js
```

Or with explicit paths:

```bash
node generate-fleet-report.js --config ./fleet-config.js --output ./output
```

The `.docx` file is written to the `output/` directory.

---

## Report Structure

| Section | Content | CAME Reference |
|---|---|---|
| Cover | Period, fleet, prepared-by, report ref | — |
| 1 | Fleet Executive Summary — status table, critical items, health indicator | §5.0 |
| 2 | Fleet Identification — aircraft register, engine/APU SNs, oil limits | §5.0.1 |
| 3 | Fleet Utilisation — FH/cycles, routes, reliability parameters, charts | §5.2.4 |
| 4 | Fleet Defect Analysis — ATA frequency, Alert Level computation, Alert Notices | §5.2.7 / §5.2.8 |
| 5 | Deferred Defect (MEL) Status — fleet summary + per-aircraft tables | §5.2.5 |
| 6 | Component Reliability (MTBUR) | §5.3.4 |
| 7 | Oil Consumption — per-engine, per-aircraft, fleet summary | §5.3.4.2 |
| 8 | Fuel Performance | §5.3.4.1 |
| 9 | Regulatory Compliance Checklist | §5.0 |
| 10 | RCB / TRB Escalation Status | §5.0.6 / §5.0.7 |
| 11 | Recommendations & Action Plan | §5.4 |
| 12 | Data Sources & Lineage | §5.3 |
| 13 | Sign-Off | §5.0.2 |

---

## Alert Level Computation

Alert Levels are calculated per **CAME §5.2.7** using a 3-sigma control chart method:

- **Baseline**: all historical months in TLP files (months not in scope are used as baseline).
- **AL = mean + 3 × standard deviation** of monthly event rates.
- **Per-aircraft AL** and **fleet-combined AL** are computed independently.
- Any ATA chapter where the current rate exceeds its AL generates an **Alert Notice** per §5.2.11.
- Alert Notices include a "Notice Due By" date = trigger date + 10 working days.

---

## Chart Generation

Charts are generated via the [QuickChart.io](https://quickchart.io) API and embedded as PNG images in the document. Set `enableCharts: false` in `fleet-config.js` to skip chart calls (useful offline or for faster iteration).

Charts produced:
- 3.1 Fleet FH per month (grouped bar)
- 3.2 Fleet Cycles per month (grouped bar)
- 3.5 Dispatch Reliability % (horizontal bar, colour-coded)
- 4.2 ATA event count per aircraft (horizontal bar)
- 7.1 Engine oil consumption with limit reference lines
- 8.1 Monthly fuel burn rate per aircraft

---

## Colour Coding

| Colour | Meaning |
|---|---|
| 🟢 GREEN | Serviceable / compliant / within limits |
| 🟡 AMBER | Warning / ground period / data incomplete |
| 🔴 RED | Alert Level exceeded / AOG / non-compliant / overdue |

---

## Post-Generation Checklist

After the document is created, the script prints a per-aircraft checklist showing which CAME data elements were populated, which were missing, and how many Alert Notices are required.

---

## Adding More Aircraft

Add additional entries to the `aircraft` array in `fleet-config.js`. Up to 7 aircraft have distinct assigned chart colours; beyond 7, colours are cycled.

---

## Troubleshooting

| Problem | Cause / Fix |
|---|---|
| `File not found` warning | The file path in `fleet-config.js` does not exist. Set to `null` if the file is not available. |
| `Not detected` in engine profile | The TLP/Fuel-Oil header rows do not contain labelled engine SN cells. Add `"Engine 1 SN:"` labels to rows 1–5 of the fuel/oil file, or set `engines:` manually in config. |
| Charts not embedded | QuickChart.io unreachable. Set `enableCharts: false` or check network access. |
| Column data missing | Column headers in your Excel files don't match expected patterns. Rename headers or adjust `parsers.js` regex patterns. |
