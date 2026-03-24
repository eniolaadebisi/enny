/**
 * ATTAIR CAMO Fleet Reliability Report — Configuration
 * =====================================================
 * STEP 1: Edit this file before running generate-fleet-report.js
 *
 * CAME Reference: ATTAIR Part 5 RCP, Issue 2 Rev 0, 30/08/2024
 */

"use strict";

module.exports = {
  // ──────────────────────────────────────────────────────────────
  // REPORTING PERIOD
  // ──────────────────────────────────────────────────────────────
  period: {
    start: "01 January 2026",       // e.g. "01 January 2026"
    end:   "31 March 2026",         // e.g. "31 March 2026"
    // Months in scope — used to filter TLP data
    // Format examples:
    //   Single month  : ["March 2026"]
    //   Quarter       : ["January 2026", "February 2026", "March 2026"]
    //   Half year     : ["January 2026","February 2026","March 2026","April 2026","May 2026","June 2026"]
    monthsInScope: ["January 2026", "February 2026", "March 2026"],
    // Human-readable label printed on the report cover
    label: "Q1 2026 — January, February, March 2026",
  },

  // ──────────────────────────────────────────────────────────────
  // FLEET IDENTIFICATION
  // ──────────────────────────────────────────────────────────────
  operator:     "ATTAIR Aviation",
  preparedBy:   "Adebisi Eniola, CAMO Reliability Officer",
  cameRef:      "ATTAIR Part 5 RCP, Issue 2 Rev 0, 30/08/2024",
  reportRef:    "ATTAIR-CAMO-FLEET-REL-2026-Q1",  // override or leave auto-generated

  // ──────────────────────────────────────────────────────────────
  // AIRCRAFT REGISTER
  // Add / remove entries for each aircraft in the fleet.
  // ──────────────────────────────────────────────────────────────
  aircraft: [
    {
      registration: "5N-BZN",
      msn:           "145556",
      type:          "Embraer ERJ-145",
      variant:       "EMB-145MP",
      engines:       "auto-detect",  // or e.g. "2 × Rolls-Royce AE3007A1P"
      apu:           "auto-detect",  // "fitted" | "not fitted" | "auto-detect"
      note:          "",             // e.g. "In heavy check Feb 2026"

      // Files for this aircraft — set path to null if not available
      files: {
        defectRecords:    "./data/5N-BZN/defect-records.xlsx",
        deferredDefects:  "./data/5N-BZN/mel-sheet.xlsx",
        tlp:              "./data/5N-BZN/tlp.xlsx",
        fuelOil:          "./data/5N-BZN/fuel-oil-log.xlsx",
        delayCancellation:"./data/5N-BZN/delay-cancel.xlsx",   // optional
        pirepMarep:       "./data/5N-BZN/pirep-marep.xlsx",    // optional
      },

      // Known data gaps (informational — not compliance failures)
      dataGaps: [
        // e.g. { item: "Fuel & Oil", month: "January 2026", reason: "Not recorded" }
      ],
    },

    {
      registration: "5N-BZQ",
      msn:           "145789",
      type:          "Embraer ERJ-145",
      variant:       "EMB-145LR",
      engines:       "auto-detect",
      apu:           "auto-detect",
      note:          "",

      files: {
        defectRecords:    "./data/5N-BZQ/defect-records.xlsx",
        deferredDefects:  "./data/5N-BZQ/mel-sheet.xlsx",
        tlp:              "./data/5N-BZQ/tlp.xlsx",
        fuelOil:          "./data/5N-BZQ/fuel-oil-log.xlsx",
        delayCancellation: null,
        pirepMarep:        null,
      },

      dataGaps: [],
    },

    // ── Add more aircraft here ────────────────────────────────────
    // {
    //   registration: "5N-BZR",
    //   msn:          "145900",
    //   type:         "Embraer ERJ-145",
    //   variant:      "EMB-145ER",
    //   engines:      "auto-detect",
    //   apu:          "auto-detect",
    //   note:         "",
    //   files: {
    //     defectRecords:    "./data/5N-BZR/defect-records.xlsx",
    //     deferredDefects:  "./data/5N-BZR/mel-sheet.xlsx",
    //     tlp:              "./data/5N-BZR/tlp.xlsx",
    //     fuelOil:          "./data/5N-BZR/fuel-oil-log.xlsx",
    //     delayCancellation: null,
    //     pirepMarep:        null,
    //   },
    //   dataGaps: [],
    // },
  ],

  // ──────────────────────────────────────────────────────────────
  // OUTPUT
  // ──────────────────────────────────────────────────────────────
  outputDir: "./output",

  // ──────────────────────────────────────────────────────────────
  // CHART GENERATION
  // Set to false to skip chart API calls (faster, offline use)
  // ──────────────────────────────────────────────────────────────
  enableCharts: true,
};
