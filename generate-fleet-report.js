/**
 * generate-fleet-report.js
 * ========================
 * ATTAIR CAMO Fleet Reliability Report Generator
 * CAME Part 5 — Reliability Control Program, Issue 2 Rev 0, 30/08/2024
 *
 * Usage:
 *   node generate-fleet-report.js [--config ./fleet-config.js] [--output ./output]
 *
 * Produces: one .docx file covering all aircraft in fleet-config.js
 */

"use strict";

const fs      = require("fs");
const path    = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, AlignmentType, HeadingLevel,
  ImageRun, PageBreak, Header, Footer, PageNumber, SimpleField,
  convertInchesToTwip, ShadingType, VerticalAlign,
  TableLayoutType,
} = require("docx");

const fetch   = require("node-fetch");
const {
  parseTLP, parseDefectRecords, parseMEL,
  parseFuelOil, parseDelayCancellation, parsePirepMarep,
  monthLabel, normaliseATA, ataName, ATA_NAMES,
} = require("./parsers");

// ─── CONFIG ──────────────────────────────────────────────────────────────────

const args   = process.argv.slice(2);
const cfgArg = args.indexOf("--config");
const cfgPath = cfgArg >= 0 ? args[cfgArg + 1] : "./fleet-config.js";
const config = require(path.resolve(cfgPath));

const OUT_DIR = (() => {
  const a = args.indexOf("--output");
  return a >= 0 ? args[a + 1] : (config.outputDir || "./output");
})();

if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

// ─── COLOUR PALETTE ──────────────────────────────────────────────────────────

const AIRCRAFT_COLOURS = [
  "1B3A6B", // Navy
  "1F6B8C", // Teal
  "D4A017", // Gold
  "1E7D34", // Green
  "6B4C9A", // Purple
  "E36C09", // Orange
];
function acColour(idx) {
  return AIRCRAFT_COLOURS[idx % AIRCRAFT_COLOURS.length];
}

const C = {
  navy:     "1B3A6B",
  teal:     "1F6B8C",
  gold:     "D4A017",
  red:      "C00000",
  amber:    "FF8C00",
  green:    "1E7D34",
  white:    "FFFFFF",
  lightBg:  "F4F6F9",
  alertBg:  "FFE8E8",
};

// ─── DOCX STYLE HELPERS ──────────────────────────────────────────────────────

function cell(text, opts = {}) {
  const {
    bold = false, italic = false, fontSize = 18,
    color = "000000", bgColor = null,
    align = AlignmentType.LEFT, colspan = 1, rowspan = 1,
    vertAlign = VerticalAlign.CENTER,
  } = opts;

  return new TableCell({
    columnSpan: colspan,
    rowSpan:    rowspan,
    verticalAlign: vertAlign,
    shading: bgColor ? { type: ShadingType.SOLID, color: bgColor, fill: bgColor } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [
      new Paragraph({
        alignment: align,
        children: [
          new TextRun({
            text:     String(text ?? ""),
            bold,
            italics:  italic,
            size:     fontSize,
            color,
          }),
        ],
      }),
    ],
  });
}

function headerCell(text, opts = {}) {
  return cell(text, { bold: true, color: C.white, bgColor: C.teal, fontSize: 18, ...opts });
}

function sectionHeading(title, cameRef = "") {
  return new Paragraph({
    children: [
      new TextRun({ text: title, bold: true, size: 28, color: C.white }),
      cameRef ? new TextRun({ text: `  ${cameRef}`, bold: false, size: 22, color: C.gold }) : null,
    ].filter(Boolean),
    shading: { type: ShadingType.SOLID, color: C.navy, fill: C.navy },
    spacing: { before: 300, after: 100 },
    border: {
      bottom: { style: BorderStyle.SINGLE, size: 6, color: C.gold },
    },
  });
}

function subHeading(title, bgColor = C.teal) {
  return new Paragraph({
    children: [new TextRun({ text: title, bold: true, size: 22, color: C.white })],
    shading: { type: ShadingType.SOLID, color: bgColor, fill: bgColor },
    spacing: { before: 200, after: 80 },
  });
}

function para(text, opts = {}) {
  const { bold = false, italic = false, size = 18, color = "000000", spacing = {} } = opts;
  return new Paragraph({
    spacing: { before: 60, after: 60, ...spacing },
    children: [new TextRun({ text: String(text ?? ""), bold, italics: italic, size, color })],
  });
}

function citationBlock(lines) {
  return new Paragraph({
    shading: { type: ShadingType.SOLID, color: C.lightBg, fill: C.lightBg },
    spacing: { before: 80, after: 80 },
    children: lines.map((l, i) =>
      new TextRun({ text: (i > 0 ? "  |  " : "📂 ") + l, italics: true, size: 16, color: "555555" })
    ),
  });
}

function alertBox(lines) {
  return new Paragraph({
    border: {
      top:    { style: BorderStyle.SINGLE, size: 12, color: C.red },
      bottom: { style: BorderStyle.SINGLE, size: 12, color: C.red },
      left:   { style: BorderStyle.SINGLE, size: 12, color: C.red },
      right:  { style: BorderStyle.SINGLE, size: 12, color: C.red },
    },
    shading: { type: ShadingType.SOLID, color: C.alertBg, fill: C.alertBg },
    spacing: { before: 120, after: 120 },
    children: lines.map((l, i) =>
      new TextRun({
        text: (i === 0 ? "⚠  " : "     ") + l + "\n",
        bold: i === 0,
        size: i === 0 ? 20 : 18,
        color: i === 0 ? C.red : "000000",
      })
    ),
  });
}

function statusColor(status) {
  if (!status) return C.lightBg;
  const s = status.toUpperCase();
  if (s === "GREEN" || s === "SERVICEABLE" || s.includes("PASS") || s === "COMPLIANT") return "C6EFCE";
  if (s === "AMBER" || s.includes("PARTIAL") || s.includes("WARN")) return "FFEB9C";
  if (s === "RED" || s.includes("FAIL") || s.includes("AOG") || s.includes("EXCEED") || s === "NON-COMPLIANT") return "FFC7CE";
  return C.lightBg;
}

function simpleTable(headers, rows, opts = {}) {
  const { stripeColor = C.lightBg } = opts;
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h) => headerCell(h)),
  });

  const dataRows = rows.map((row, ri) =>
    new TableRow({
      children: row.map((v, ci) => {
        // Check if it's a status-bearing cell (last column or explicit)
        const bg = ri % 2 === 1 ? stripeColor : "FFFFFF";
        if (typeof v === "object" && v !== null && v.__status) {
          return cell(v.text, { bgColor: statusColor(v.__status), ...v.opts });
        }
        return cell(v, { bgColor: bg });
      }),
    })
  );

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    rows: [headerRow, ...dataRows],
  });
}

function statusCell(text, status, opts = {}) {
  return { __status: status, text, opts };
}

// ─── CHART HELPER ────────────────────────────────────────────────────────────

async function fetchChart(config, width = 600, height = 350) {
  if (!config.enableCharts) return null;
  try {
    const url = "https://quickchart.io/chart";
    const body = JSON.stringify({ width, height, backgroundColor: "white", chart: config });
    const res = await fetch(url, {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body,
      signal: AbortSignal.timeout(15000),
    });
    if (!res.ok) {
      console.warn(`  ⚠  Chart API error: ${res.status}`);
      return null;
    }
    const buf = Buffer.from(await res.arrayBuffer());
    return buf;
  } catch (e) {
    console.warn(`  ⚠  Chart fetch failed: ${e.message}`);
    return null;
  }
}

function chartImage(buf, w = 600, h = 350) {
  if (!buf) {
    return para("[Chart unavailable — QuickChart.io could not be reached]",
      { italic: true, color: "888888" });
  }
  return new Paragraph({
    children: [new ImageRun({ data: buf, transformation: { width: w, height: h } })],
    spacing: { before: 100, after: 100 },
  });
}

// ─── ANALYTICS HELPERS ───────────────────────────────────────────────────────

/** Build monthly ATA frequency map from defect records */
function buildATAFrequency(defects, scopeKeys) {
  const freq = {}; // { "2026-01": { "32": 3, "29": 1 } }
  for (const d of defects) {
    if (!d.date) continue;
    const k = `${d.date.getFullYear()}-${String(d.date.getMonth() + 1).padStart(2, "0")}`;
    if (!scopeKeys.includes(k)) continue;
    const ata = normaliseATA(d.ata) || "UNKNOWN";
    if (!freq[k]) freq[k] = {};
    freq[k][ata] = (freq[k][ata] || 0) + 1;
  }
  return freq;
}

/** Aggregate frequency across all scope months: { ata: totalCount } */
function totalATAFreq(freqMap, scopeKeys) {
  const total = {};
  for (const k of scopeKeys) {
    const m = freqMap[k] || {};
    for (const [ata, cnt] of Object.entries(m)) {
      total[ata] = (total[ata] || 0) + cnt;
    }
  }
  return total;
}

/**
 * Compute Alert Level (AL) per ATA chapter using 3-sigma control chart method.
 * Uses ALL available historical months (not just scope months) as baseline.
 *
 * Returns: { ata: { mean, stddev, al, scopeRate, status } }
 * Status: "GREEN" | "AMBER" | "RED"
 */
function computeAlertLevels(freqMap, allMonthKeys, scopeKeys, totalFH) {
  // Collect monthly rates per ATA (events / 1000 FH if FH available, else raw count)
  const ataMonthly = {}; // { ata: [rate per historical month] }

  // Build the set of months that provide baseline (historical + scope)
  const historicalKeys = allMonthKeys.filter((k) => !scopeKeys.includes(k));
  const baselineKeys = historicalKeys.length >= 3 ? historicalKeys : allMonthKeys;

  for (const k of baselineKeys) {
    const m = freqMap[k] || {};
    for (const [ata, cnt] of Object.entries(m)) {
      if (!ataMonthly[ata]) ataMonthly[ata] = [];
      ataMonthly[ata].push(cnt);
    }
  }

  const results = {};
  const allATAs = new Set([
    ...Object.keys(ataMonthly),
    ...scopeKeys.flatMap((k) => Object.keys(freqMap[k] || {})),
  ]);

  for (const ata of allATAs) {
    const series = ataMonthly[ata] || [];
    // Pad with zeros for months with no events (included in baseline)
    const paddedLen = Math.max(series.length, baselineKeys.length, 3);
    const padded = [...series, ...Array(paddedLen - series.length).fill(0)];

    const mean   = padded.reduce((a, b) => a + b, 0) / padded.length;
    const variance = padded.reduce((s, x) => s + (x - mean) ** 2, 0) / padded.length;
    const stddev = Math.sqrt(variance);
    const al     = +(mean + 3 * stddev).toFixed(2);

    // Current scope rate = sum of events in scope months
    const scopeTotal = scopeKeys.reduce((s, k) => s + ((freqMap[k] || {})[ata] || 0), 0);
    const scopeMonthCount = scopeKeys.length || 1;
    const scopeRate = +(scopeTotal / scopeMonthCount).toFixed(2);

    let status = "GREEN";
    if (scopeRate > al)          status = "RED";
    else if (scopeRate > al * 0.8) status = "AMBER";

    results[ata] = { mean: +mean.toFixed(2), stddev: +stddev.toFixed(2), al, scopeRate, status, scopeTotal };
  }

  return results;
}

/** Build top routes from TLP sectors */
function topRoutes(months, topN = 10) {
  const routes = {};
  for (const m of months) {
    for (const s of m.sectors || []) {
      if (!s.from || !s.to) continue;
      const key = `${s.from}-${s.to}`;
      if (!routes[key]) routes[key] = { from: s.from, to: s.to, count: 0, totalFH: 0 };
      routes[key].count++;
      if (s.sectorFH) routes[key].totalFH += s.sectorFH;
    }
  }
  return Object.values(routes)
    .sort((a, b) => b.count - a.count)
    .slice(0, topN)
    .map((r) => ({ ...r, avgFH: r.count ? +(r.totalFH / r.count).toFixed(2) : null }));
}

/** Compute dispatch reliability % */
function dispatchReliability(sectors, delays, cancellations) {
  const totalOps = sectors + delays + cancellations;
  if (!totalOps) return null;
  return +((sectors / totalOps) * 100).toFixed(2);
}

/** MTBUR calculation: FH / removals */
function mtbur(fh, removals) {
  if (!fh || !removals) return null;
  return +(fh / removals).toFixed(1);
}

/** Format a number to 2 decimal places or return "N/A" */
function fmt(v, decimals = 2) {
  if (v == null) return "N/A";
  return Number(v).toFixed(decimals);
}

/** Format a date as DD MMM YYYY */
function fmtDate(d) {
  if (!d) return "N/A";
  return d.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

/** Add N working days to a date (Mon-Fri) */
function addWorkingDays(date, n) {
  const d = new Date(date);
  let added = 0;
  while (added < n) {
    d.setDate(d.getDate() + 1);
    if (d.getDay() !== 0 && d.getDay() !== 6) added++;
  }
  return d;
}

// ─── MAIN ─────────────────────────────────────────────────────────────────────

async function main() {
  console.log("\n╔══════════════════════════════════════════════════════════╗");
  console.log("║  ATTAIR Fleet Reliability Report Generator               ║");
  console.log("║  CAME Part 5 RCP — Issue 2 Rev 0 — 30/08/2024           ║");
  console.log("╚══════════════════════════════════════════════════════════╝\n");

  const { period, operator, preparedBy, cameRef, aircraft } = config;
  const scopeKeys = period.monthsInScope.map(monthLabel);

  // Auto-generate report ref if not set
  const reportRef = config.reportRef ||
    `ATTAIR-CAMO-FLEET-REL-${new Date().getFullYear()}-${
      String(new Date().getMonth() + 1).padStart(2, "0")}`;

  // ── STEP: Parse all aircraft files independently ──────────────────────────
  console.log("► Parsing aircraft data…");
  const acData = [];

  for (const ac of aircraft) {
    console.log(`  [${ac.registration}] Loading files…`);
    const files = ac.files || {};

    const [tlp, defects, mel, fuelOil, delays, pirep] = await Promise.all([
      parseTLP(files.tlp, ac.registration, period.monthsInScope),
      parseDefectRecords(files.defectRecords, ac.registration),
      parseMEL(files.deferredDefects, ac.registration),
      parseFuelOil(files.fuelOil, ac.registration),
      parseDelayCancellation(files.delayCancellation, ac.registration),
      parsePirepMarep(files.pirepMarep, ac.registration),
    ]);

    // Merge engine profile: TLP takes priority, then FuelOil
    const ep = {
      eng1SN:   tlp?.engineProfile?.eng1SN   || fuelOil?.engineProfile?.eng1SN   || "Not detected",
      eng2SN:   tlp?.engineProfile?.eng2SN   || fuelOil?.engineProfile?.eng2SN   || "Not detected",
      oilLimit: tlp?.engineProfile?.oilLimit || fuelOil?.engineProfile?.oilLimit || "Not stated — verify against FIM",
      apuSN:    tlp?.engineProfile?.apuSN    || fuelOil?.engineProfile?.apuSN    || (ac.apu === "not fitted" ? "Not fitted" : "Not detected"),
    };
    const engType = ac.engines === "auto-detect"
      ? (ep.eng1SN !== "Not detected" ? `Detected SN: ${ep.eng1SN}` : "Auto-detect — not found in files")
      : ac.engines;

    // Scope months TLP data
    const scopeMonths = (tlp?.months || []).filter((m) => m.scopeMonth);
    const allMonths   = tlp?.months || [];
    const allMonthKeys = allMonths.map((m) => m.ymKey);

    // Utilisation totals for scope
    const totalFH  = scopeMonths.reduce((s, m) => s + (m.fh || 0), 0);
    const totalCYC = scopeMonths.reduce((s, m) => s + (m.cyc || 0), 0);

    // Defect frequency & alert levels
    const defFreq    = buildATAFrequency(defects, scopeKeys);
    const totalFreq  = totalATAFreq(defFreq, scopeKeys);
    const alertLevels = computeAlertLevels(defFreq, allMonthKeys, scopeKeys, totalFH);

    // MEL stats
    const openMELs     = mel.filter((m) => m.isOpen);
    const closedPeriod = mel.filter((m) => {
      if (!m.closedDate) return false;
      const k = `${m.closedDate.getFullYear()}-${String(m.closedDate.getMonth() + 1).padStart(2, "0")}`;
      return scopeKeys.includes(k);
    });

    // Fuel/oil scope data
    const fuelMonths = (fuelOil?.months || []).filter(
      (m) => scopeKeys.includes(m.ymKey)
    );

    // Delay stats
    const scopeDelays = delays.filter((d) => {
      if (!d.date) return false;
      const k = `${d.date.getFullYear()}-${String(d.date.getMonth() + 1).padStart(2, "0")}`;
      return scopeKeys.includes(k);
    });
    const techDelays  = scopeDelays.filter((d) => d.type === "DELAY");
    const cancels     = scopeDelays.filter((d) => d.type === "CANCELLATION");
    const totalSectors = scopeMonths.reduce((s, m) => s + m.sectors.length, 0);
    const dispRel = dispatchReliability(totalSectors - techDelays.length - cancels.length,
      techDelays.length, cancels.length);

    // PIREP/MAREP
    const scopePirep = pirep.filter((p) => {
      if (!p.date) return false;
      const k = `${p.date.getFullYear()}-${String(p.date.getMonth() + 1).padStart(2, "0")}`;
      return scopeKeys.includes(k);
    });

    // Overall status
    const hasALBreach = Object.values(alertLevels).some((v) => v.status === "RED");
    const hasGroundPeriod = scopeMonths.some((m) => m.isGroundPeriod);
    const overallStatus = totalFH === 0 ? "GROUND PERIOD"
      : hasALBreach ? "⚠ AL EXCEEDANCE"
      : "SERVICEABLE";

    acData.push({
      reg:           ac.registration,
      msn:           ac.msn,
      type:          ac.type,
      variant:       ac.variant,
      note:          ac.note,
      files:         ac.files,
      dataGaps:      ac.dataGaps || [],
      engineProfile: ep,
      engType,
      // raw parsed data
      tlpData:       tlp,
      defects,
      mel,
      fuelOil,
      delays,
      pirep,
      // derived
      scopeMonths,
      allMonths,
      allMonthKeys,
      totalFH:       +totalFH.toFixed(2),
      totalCYC,
      defFreq,
      totalFreq,
      alertLevels,
      openMELs,
      closedPeriod,
      fuelMonths,
      scopeDelays,
      techDelays,
      cancels,
      totalSectors,
      dispRel,
      scopePirep,
      hasALBreach,
      hasGroundPeriod,
      overallStatus,
    });

    console.log(`  [${ac.registration}] ✓  FH:${totalFH.toFixed(1)}  CYC:${totalCYC}  Defects:${defects.length}  Open MELs:${openMELs.length}`);
  }

  // ── Fleet-level aggregates ────────────────────────────────────────────────
  const fleetTotalFH  = +acData.reduce((s, a) => s + a.totalFH, 0).toFixed(2);
  const fleetTotalCYC = acData.reduce((s, a) => s + a.totalCYC, 0);
  const fleetOpenMELs = acData.reduce((s, a) => s + a.openMELs.length, 0);

  // Fleet ATA totals (sum across all aircraft)
  const fleetATATotals = {};
  for (const ac of acData) {
    for (const [ata, cnt] of Object.entries(ac.totalFreq)) {
      fleetATATotals[ata] = (fleetATATotals[ata] || 0) + cnt;
    }
  }

  // Fleet-level alert levels (pool all aircraft defect frequency maps)
  const pooledFreq = {};
  for (const ac of acData) {
    for (const [k, ataMap] of Object.entries(ac.defFreq)) {
      if (!pooledFreq[k]) pooledFreq[k] = {};
      for (const [ata, cnt] of Object.entries(ataMap)) {
        pooledFreq[k][ata] = (pooledFreq[k][ata] || 0) + cnt;
      }
    }
  }
  const fleetAllMonthKeys = [...new Set(acData.flatMap((a) => a.allMonthKeys))].sort();
  const fleetALs = computeAlertLevels(pooledFreq, fleetAllMonthKeys, scopeKeys, fleetTotalFH);

  console.log("\n► Building document…");

  // ════════════════════════════════════════════════════════════════════════════
  // DOCUMENT SECTIONS
  // ════════════════════════════════════════════════════════════════════════════

  const children = [];

  // ── COVER PAGE ─────────────────────────────────────────────────────────────

  children.push(
    new Paragraph({
      children: [new TextRun({ text: "ATTAIR AVIATION", bold: true, size: 56, color: C.navy })],
      alignment: AlignmentType.CENTER,
      spacing: { before: 1440, after: 200 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "Fleet Reliability Report", bold: true, size: 40, color: C.teal })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
    }),
    new Paragraph({
      children: [new TextRun({ text: period.label, bold: true, size: 28, color: C.gold })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
    }),
    new Paragraph({
      children: [new TextRun({ text: reportRef, size: 22, color: "555555" })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
    }),
    new Paragraph({
      children: [new TextRun({ text: `Prepared by: ${preparedBy}`, size: 20 })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
    }),
    new Paragraph({
      children: [new TextRun({ text: cameRef, size: 18, italics: true, color: "555555" })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
    }),
    new Paragraph({
      children: [new TextRun({ text: `Generated: ${new Date().toLocaleDateString("en-GB", { day: "2-digit", month: "long", year: "numeric" })}`, size: 18, color: "888888" })],
      alignment: AlignmentType.CENTER,
      spacing: { after: 1440 },
    }),
    new Paragraph({ children: [new PageBreak()] })
  );

  // ═══════════════════════════════════════════
  // SECTION 1 — FLEET EXECUTIVE SUMMARY
  // ═══════════════════════════════════════════

  children.push(sectionHeading("SECTION 1 — FLEET EXECUTIVE SUMMARY", "CAME §5.0"));

  // Fleet status table
  children.push(subHeading("1.1 Fleet Status Overview"));
  children.push(
    simpleTable(
      ["Reg", "Type", "Variant", "MSN", "Period FH", "Period CYC",
       "Open MELs", "Oil Status", "Dispatch Rel%", "Ground Periods", "Overall Status"],
      acData.map((ac) => {
        const oilOk = ac.fuelMonths.every((m) =>
          (m.eng1OilRateLHr == null || parseFloat(ac.engineProfile.oilLimit) == null ||
           m.eng1OilRateLHr <= parseFloat(ac.engineProfile.oilLimit)) &&
          (m.eng2OilRateLHr == null || m.eng2OilRateLHr <= parseFloat(ac.engineProfile.oilLimit))
        );
        const oilStatus = ac.fuelMonths.length === 0 ? "No data" : (oilOk ? "NORMAL" : "EXCEED");
        const gpCount = ac.scopeMonths.filter((m) => m.isGroundPeriod).length;
        const statusCol = ac.overallStatus.includes("AL") ? "RED"
          : ac.overallStatus === "GROUND PERIOD" ? "AMBER" : "GREEN";

        return [
          ac.reg, ac.type, ac.variant, ac.msn,
          fmt(ac.totalFH, 1), ac.totalCYC,
          ac.openMELs.length,
          statusCell(oilStatus, oilOk ? "GREEN" : "RED"),
          statusCell(
            ac.dispRel != null ? `${fmt(ac.dispRel)}%` : "N/A",
            ac.dispRel == null ? "AMBER"
              : ac.dispRel >= 98.5 ? "GREEN"
              : ac.dispRel >= 95 ? "AMBER" : "RED"
          ),
          gpCount > 0 ? `${gpCount} month(s)` : "None",
          statusCell(ac.overallStatus, statusCol),
        ];
      })
    )
  );

  // Fleet totals row
  children.push(
    para(`Fleet Totals — Total FH: ${fmt(fleetTotalFH, 1)} | Total Cycles: ${fleetTotalCYC} | Combined Open MELs: ${fleetOpenMELs}`,
      { bold: true, size: 18 })
  );

  // Critical items panel
  const redAlerts = [];
  for (const ac of acData) {
    for (const [ata, alv] of Object.entries(ac.alertLevels)) {
      if (alv.status === "RED") {
        redAlerts.push({
          reg: ac.reg,
          ata,
          al: alv.al,
          scopeRate: alv.scopeRate,
          due: fmtDate(addWorkingDays(new Date(), 10)),
        });
      }
    }
  }

  if (redAlerts.length > 0) {
    children.push(subHeading("1.2 Critical Items — Alert Level Exceedances", C.red));
    for (const a of redAlerts) {
      children.push(alertBox([
        `ALERT NOTICE — ${a.reg} — CAME §5.2.11`,
        `ATA: ${a.ata} | Alert Level: ${a.al} | Current Rate: ${a.scopeRate}`,
        `Notice Due By: ${a.due}`,
        `Action: Formal investigation — report to RCB`,
      ]));
    }
  }

  // Fleet health indicator
  children.push(subHeading("1.3 Fleet Health Indicator"));
  children.push(
    simpleTable(
      ["Registration", "AL Status", "Open MELs", "Oil Status", "Health"],
      acData.map((ac) => {
        const al = ac.hasALBreach ? "RED — Exceedance" : "GREEN";
        const health = ac.hasALBreach ? "RED" : ac.openMELs.length > 0 ? "AMBER" : "GREEN";
        return [
          ac.reg,
          statusCell(al, ac.hasALBreach ? "RED" : "GREEN"),
          ac.openMELs.length,
          ac.fuelMonths.length === 0 ? "AMBER — No data" : "GREEN",
          statusCell(health, health),
        ];
      })
    )
  );

  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ═══════════════════════════════════════════
  // SECTION 2 — FLEET IDENTIFICATION
  // ═══════════════════════════════════════════

  children.push(sectionHeading("SECTION 2 — FLEET IDENTIFICATION", "CAME §5.0.1"));

  children.push(
    simpleTable(
      ["No.", "Reg", "Type", "Variant", "MSN", "Engine 1 SN", "Engine 2 SN",
       "APU SN", "Engine Type", "Oil Limit"],
      acData.map((ac, i) => [
        i + 1, ac.reg, ac.type, ac.variant, ac.msn,
        ac.engineProfile.eng1SN,
        ac.engineProfile.eng2SN,
        ac.engineProfile.apuSN,
        ac.engType,
        ac.engineProfile.oilLimit,
      ])
    )
  );

  // Flag variant / oil limit differences
  const uniqueLimits = [...new Set(acData.map((a) => a.engineProfile.oilLimit))];
  if (uniqueLimits.length > 1) {
    children.push(
      para("⚠  Oil limits differ between aircraft variants — limits applied independently per aircraft in Section 7.",
        { bold: true, color: C.red })
    );
  }
  for (const acNote of acData.filter((a) => a.note)) {
    children.push(para(`Note [${acNote.reg}]: ${acNote.note}`, { italic: true }));
  }

  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ═══════════════════════════════════════════
  // SECTION 3 — FLEET UTILISATION
  // ═══════════════════════════════════════════

  children.push(sectionHeading("SECTION 3 — FLEET UTILISATION", "CAME §5.2.4"));

  // Fleet summary table
  children.push(subHeading("3.1 Fleet Monthly Summary"));
  const monthRows = scopeKeys.map((k) => {
    const label = period.monthsInScope[scopeKeys.indexOf(k)];
    let fleetFH = 0, fleetCYC = 0, inService = 0, onGround = 0;
    for (const ac of acData) {
      const m = ac.scopeMonths.find((mo) => mo.ymKey === k);
      if (m) {
        fleetFH  += m.fh  || 0;
        fleetCYC += m.cyc || 0;
        if (m.isGroundPeriod) onGround++; else inService++;
      } else {
        onGround++;
      }
    }
    const totalSecs = acData.reduce((s, ac) => {
      const m = ac.scopeMonths.find((mo) => mo.ymKey === k);
      return s + (m ? m.sectors.length : 0);
    }, 0);
    const avgSec = fleetCYC > 0 ? +(fleetFH / fleetCYC).toFixed(2) : null;
    return [label, fmt(fleetFH, 1), fleetCYC, inService, onGround, fmt(avgSec, 2)];
  });
  // Totals row
  monthRows.push([
    "TOTAL",
    fmt(fleetTotalFH, 1),
    fleetTotalCYC,
    "",
    "",
    "",
  ]);
  children.push(simpleTable(
    ["Month", "Fleet FH", "Fleet Cycles", "Aircraft In-Service", "Aircraft on Ground", "Fleet Avg Sector (hrs)"],
    monthRows
  ));

  // Per-aircraft utilisation
  children.push(subHeading("3.2 Per-Aircraft Utilisation"));
  const utilRows = [];
  for (const ac of acData) {
    for (const k of scopeKeys) {
      const m = ac.scopeMonths.find((mo) => mo.ymKey === k);
      const label = period.monthsInScope[scopeKeys.indexOf(k)];
      if (!m) {
        utilRows.push([ac.reg, label, "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A", "N/A",
          statusCell("NO DATA", "AMBER")]);
        continue;
      }
      const avgSec = m.cyc ? +(m.fh / m.cyc).toFixed(2) : null;
      const status = m.isGroundPeriod ? "GROUND PERIOD" : "SERVICEABLE";
      utilRows.push([
        ac.reg, label,
        m.operatingDays, m.nilOpDays,
        fmt(m.fh, 1), m.cyc,
        fmt(avgSec, 2),
        fmt(m.closingTSN, 1), m.closingCSN,
        fmt(m.eng1FH, 1), fmt(m.eng2FH, 1),
        fmt(m.apuHrs, 1), m.apuStarts ?? "N/A",
        statusCell(status, m.isGroundPeriod ? "AMBER" : "GREEN"),
      ]);
    }
  }
  children.push(simpleTable(
    ["Reg", "Month", "Op Days", "Nil-Op", "FH", "CYC",
     "Avg Sector", "AF TSN", "AF CSN",
     "Eng1 FH", "Eng2 FH", "APU Hrs", "APU Starts", "Status"],
    utilRows
  ));

  // Top routes per aircraft
  children.push(subHeading("3.3 Top Routes — Fleet Wide"));
  const allRouteRows = [];
  for (const ac of acData) {
    const routes = topRoutes(ac.scopeMonths, 5);
    for (const r of routes) {
      allRouteRows.push([ac.reg, `${r.from} — ${r.to}`, r.count, fmt(r.avgFH, 2)]);
    }
  }
  if (allRouteRows.length > 0) {
    children.push(simpleTable(
      ["Aircraft", "Route", "Frequency", "Avg Sector (hrs)"],
      allRouteRows
    ));
  } else {
    children.push(para("No route data extracted from TLP files.", { italic: true }));
  }

  // Citation blocks
  for (const ac of acData) {
    children.push(citationBlock([
      `Aircraft: ${ac.reg}`,
      `File: ${ac.files?.tlp || "N/A"}`,
      `Sheets: ${ac.tlpData?.rawSheetNames?.join(", ") || "N/A"}`,
      `CAME type: Operational Data`,
    ]));
  }

  // Reliability parameters
  children.push(subHeading("3.4 Reliability Parameters (CAME §5.2.4)"));
  const paramRows = [];
  for (const ac of acData) {
    const calDaysTotal = ac.scopeMonths.reduce((s, m) => s + m.calDays, 0) || 1;
    const dailyUtilHrs = +(ac.totalFH / calDaysTotal).toFixed(2);
    const dailyUtilCyc = +(ac.totalCYC / calDaysTotal).toFixed(2);
    const avgFL = ac.totalCYC ? +(ac.totalFH / ac.totalCYC).toFixed(2) : null;
    const techDelRate = ac.totalSectors
      ? +((ac.techDelays.length / ac.totalSectors) * 100).toFixed(2) : null;
    const cancelRate = ac.totalSectors
      ? +((ac.cancels.length / ac.totalSectors) * 100).toFixed(2) : null;
    const pirepRate = ac.totalFH
      ? +((ac.scopePirep.filter((p) => p.type === "PIREP").length / ac.totalFH) * 100).toFixed(2) : null;
    const marepRate = ac.totalFH
      ? +((ac.scopePirep.filter((p) => p.type === "MAREP").length / ac.totalFH) * 100).toFixed(2) : null;

    paramRows.push([
      ac.reg,
      `${fmt(dailyUtilHrs)} hrs`,
      `${fmt(dailyUtilCyc)} cyc`,
      fmt(avgFL),
      techDelRate != null ? `${fmt(techDelRate)}%` : "N/A",
      cancelRate  != null ? `${fmt(cancelRate)}%`  : "N/A",
      statusCell(
        ac.dispRel != null ? `${fmt(ac.dispRel)}%` : "N/A",
        ac.dispRel == null ? "AMBER" : ac.dispRel >= 98.5 ? "GREEN" : ac.dispRel >= 95 ? "AMBER" : "RED"
      ),
      pirepRate != null ? fmt(pirepRate) : "N/A",
      marepRate != null ? fmt(marepRate) : "N/A",
    ]);
  }
  children.push(simpleTable(
    ["Reg", "Daily Util (hrs)", "Daily Util (cyc)", "Avg FL (hrs)",
     "Tech Delay Rate%", "Cancel Rate%", "Dispatch Rel%", "PIREP Rate", "MAREP Rate"],
    paramRows
  ));

  // Charts 3.1–3.7
  if (config.enableCharts) {
    children.push(subHeading("3.5 Utilisation Charts"));

    // Chart 3.1 — Fleet FH per month per aircraft
    const c31 = {
      type: "bar",
      data: {
        labels: period.monthsInScope,
        datasets: acData.map((ac, i) => ({
          label: ac.reg,
          data: scopeKeys.map((k) => {
            const m = ac.scopeMonths.find((mo) => mo.ymKey === k);
            return m ? (m.fh || 0) : 0;
          }),
          backgroundColor: `#${acColour(i)}`,
        })),
      },
      options: {
        plugins: { title: { display: true, text: "Chart 3.1 — Fleet FH per Month [CAME §5.2.4-4a]" } },
        scales: { y: { beginAtZero: true, title: { display: true, text: "Flight Hours" } } },
      },
    };
    const img31 = await fetchChart({ ...config, chart: c31 });
    children.push(chartImage(img31));

    // Chart 3.2 — Fleet Cycles per month
    const c32 = {
      type: "bar",
      data: {
        labels: period.monthsInScope,
        datasets: acData.map((ac, i) => ({
          label: ac.reg,
          data: scopeKeys.map((k) => {
            const m = ac.scopeMonths.find((mo) => mo.ymKey === k);
            return m ? (m.cyc || 0) : 0;
          }),
          backgroundColor: `#${acColour(i)}`,
        })),
      },
      options: {
        plugins: { title: { display: true, text: "Chart 3.2 — Fleet Cycles per Month [CAME §5.2.4-4a]" } },
        scales: { y: { beginAtZero: true, title: { display: true, text: "Cycles" } } },
      },
    };
    const img32 = await fetchChart({ ...config, chart: c32 });
    children.push(chartImage(img32));

    // Chart 3.5 — Dispatch Reliability
    const c35 = {
      type: "horizontalBar",
      data: {
        labels: acData.map((a) => a.reg),
        datasets: [{
          label: "Dispatch Reliability %",
          data: acData.map((a) => a.dispRel ?? 0),
          backgroundColor: acData.map((a) =>
            a.dispRel == null ? "#888888"
              : a.dispRel >= 98.5 ? "#1E7D34"
              : a.dispRel >= 95 ? "#FF8C00" : "#C00000"
          ),
        }],
      },
      options: {
        plugins: { title: { display: true, text: "Chart 3.5 — Dispatch Reliability % [CAME §5.2.4-4d]" },
          annotation: { annotations: [{ type: "line", mode: "vertical", scaleID: "x-axis-0",
            value: 98.5, borderColor: "#C00000", borderWidth: 2, label: { content: "Target 98.5%", enabled: true } }] }
        },
        scales: { x: { min: 90, max: 100 } },
      },
    };
    const img35 = await fetchChart({ ...config, chart: c35 });
    children.push(chartImage(img35));
  }

  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ═══════════════════════════════════════════
  // SECTION 4 — FLEET DEFECT ANALYSIS
  // ═══════════════════════════════════════════

  children.push(sectionHeading("SECTION 4 — FLEET DEFECT ANALYSIS", "CAME §5.2.7 / §5.2.8"));

  // Fleet ATA frequency table
  children.push(subHeading("4.1 Fleet ATA Frequency Table"));
  const fleetATARows = Object.entries(fleetATATotals)
    .sort((a, b) => b[1] - a[1])
    .map(([ata, total]) => {
      const risk = total >= 6 ? "RED" : total >= 3 ? "AMBER" : "GREEN";
      const perAC = acData.map((ac) => ac.totalFreq[ata] || 0);
      return [
        ata,
        ataName(ata),
        ...perAC,
        total,
        statusCell(risk, risk),
      ];
    });

  if (fleetATARows.length > 0) {
    children.push(simpleTable(
      ["ATA", "System", ...acData.map((a) => a.reg), "Fleet Total", "Risk Level"],
      fleetATARows
    ));
  } else {
    children.push(para("No defect events found in scope period.", { italic: true }));
  }

  // Per-aircraft defect tables
  for (const ac of acData) {
    children.push(subHeading(`4.2 ${ac.reg} — Defect Events by ATA`, C.navy));
    const acDefRows = Object.entries(ac.totalFreq)
      .sort((a, b) => b[1] - a[1])
      .map(([ata, cnt]) => {
        const alv = ac.alertLevels[ata];
        return [
          ata, ataName(ata), cnt,
          alv ? fmt(alv.mean, 2) : "N/A",
          alv ? fmt(alv.al, 2) : "N/A",
          statusCell(alv ? alv.status : "GREEN", alv ? alv.status : "GREEN"),
        ];
      });

    if (acDefRows.length > 0) {
      children.push(simpleTable(
        ["ATA", "System", "Events (Scope)", "Hist Mean", "Alert Level", "Status"],
        acDefRows
      ));
    } else {
      children.push(para(`No defect events found for ${ac.reg}.`, { italic: true }));
    }

    children.push(citationBlock([
      `Aircraft: ${ac.reg}`,
      `File: ${ac.files?.defectRecords || "N/A"}`,
      `Records: ${ac.defects.length}`,
      `CAME type: Routine Task Finding`,
    ]));
  }

  // Combined AL table (per-aircraft + fleet)
  children.push(subHeading("4.3 Combined Alert Level Status (Per-Aircraft + Fleet) — CAME §5.2.7"));
  const allATAs = [...new Set([
    ...Object.keys(fleetATATotals),
    ...acData.flatMap((a) => Object.keys(a.alertLevels)),
  ])].sort();

  if (allATAs.length > 0) {
    const alHeaders = ["ATA", "System"];
    for (const ac of acData) alHeaders.push(`${ac.reg} Rate`, `${ac.reg} AL`, `${ac.reg} Status`);
    alHeaders.push("Fleet Rate", "Fleet AL", "Fleet Status");

    const alRows = allATAs.map((ata) => {
      const row = [ata, ataName(ata)];
      for (const ac of acData) {
        const alv = ac.alertLevels[ata];
        row.push(alv ? fmt(alv.scopeRate, 2) : "0");
        row.push(alv ? fmt(alv.al, 2) : "N/A");
        row.push(statusCell(alv ? alv.status : "GREEN", alv ? alv.status : "GREEN"));
      }
      const fal = fleetALs[ata];
      row.push(fal ? fmt(fal.scopeRate, 2) : "0");
      row.push(fal ? fmt(fal.al, 2) : "N/A");
      row.push(statusCell(fal ? fal.status : "GREEN", fal ? fal.status : "GREEN"));
      return row;
    });

    children.push(simpleTable(alHeaders, alRows));
  }

  // Alert notices for all RED findings
  const allRedFindings = [];
  for (const ac of acData) {
    for (const [ata, alv] of Object.entries(ac.alertLevels)) {
      if (alv.status === "RED") {
        allRedFindings.push({ reg: ac.reg, ata, alv, source: "per-aircraft" });
      }
    }
  }
  for (const [ata, alv] of Object.entries(fleetALs)) {
    if (alv.status === "RED") {
      allRedFindings.push({ reg: "FLEET", ata, alv, source: "fleet" });
    }
  }

  for (const f of allRedFindings) {
    const dueDate = fmtDate(addWorkingDays(new Date(), 10));
    children.push(alertBox([
      `ALERT NOTICE — ${f.reg} — CAME §5.2.11`,
      `ATA: ${f.ata} (${ataName(f.ata)}) | Alert Level: ${fmt(f.alv.al, 2)} | Current Rate: ${fmt(f.alv.scopeRate, 2)}`,
      `Trigger Date: ${new Date().toLocaleDateString("en-GB")} | Notice Due By: ${dueDate}`,
      `Send To: Approved Maintenance Organisation`,
      `Action: Formal investigation — report to RCB`,
      `Fleet Note: Check if same ATA trend exists on other aircraft of same type.`,
      `Status: [ ] Issued  [ ] Pending`,
    ]));
  }

  // Charts 4.1–4.2
  if (config.enableCharts && fleetATARows.length > 0) {
    children.push(subHeading("4.4 Defect Analysis Charts"));
    const c42Labels = Object.keys(fleetATATotals).sort((a, b) => fleetATATotals[b] - fleetATATotals[a]).slice(0, 12);
    const c42 = {
      type: "horizontalBar",
      data: {
        labels: c42Labels.map((a) => `${a} ${ataName(a)}`),
        datasets: acData.map((ac, i) => ({
          label: ac.reg,
          data: c42Labels.map((a) => ac.totalFreq[a] || 0),
          backgroundColor: `#${acColour(i)}`,
        })),
      },
      options: {
        plugins: { title: { display: true, text: "Chart 4.2 — ATA Event Count per Aircraft [CAME §5.2.7]" } },
        scales: { x: { beginAtZero: true } },
      },
    };
    const img42 = await fetchChart({ ...config, chart: c42 });
    children.push(chartImage(img42));
  }

  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ═══════════════════════════════════════════
  // SECTION 5 — DEFERRED DEFECT (MEL) STATUS
  // ═══════════════════════════════════════════

  children.push(sectionHeading("SECTION 5 — DEFERRED DEFECT (MEL) STATUS", "CAME §5.2.5"));

  // Fleet MEL summary
  children.push(subHeading("5.1 Fleet MEL Summary"));
  children.push(
    simpleTable(
      ["Reg", "Open MELs", "Closed This Period", "Cat A", "Cat B", "Cat C", "Cat D", "Any Overdue?"],
      acData.map((ac) => {
        const cats = { A: 0, B: 0, C: 0, D: 0 };
        for (const m of ac.openMELs) {
          const cat = (m.category || "").toUpperCase();
          if (cats[cat] !== undefined) cats[cat]++;
        }
        const overdue = ac.openMELs.some((m) => m.isOverdue);
        return [
          ac.reg,
          ac.openMELs.length,
          ac.closedPeriod.length,
          cats.A, cats.B, cats.C, cats.D,
          statusCell(overdue ? "YES" : "NO", overdue ? "RED" : "GREEN"),
        ];
      })
    )
  );

  // Per-aircraft MEL tables
  for (const ac of acData) {
    children.push(subHeading(`5.2 ${ac.reg} — MEL / Deferred Defect Detail`, C.navy));
    if (ac.mel.length === 0) {
      children.push(para(`No MEL items found for ${ac.reg}.`, { italic: true }));
    } else {
      const melRows = ac.mel.map((m) => [
        m.ref || "—",
        m.ata || "—",
        ataName(m.ata),
        m.description || "—",
        m.category || "—",
        fmtDate(m.raisedDate),
        fmtDate(m.dueDate),
        fmtDate(m.closedDate),
        m.daysOpen ?? "—",
        statusCell(m.status, m.isOverdue ? "RED" : m.isOpen ? "AMBER" : "GREEN"),
      ]);
      children.push(simpleTable(
        ["Ref", "ATA", "System", "Description", "Cat", "Raised", "Due", "Closed", "Days Open", "Status"],
        melRows
      ));
    }
    children.push(citationBlock([
      `Aircraft: ${ac.reg}`,
      `File: ${ac.files?.deferredDefects || "N/A"}`,
      `Records: ${ac.mel.length}`,
      `Open: ${ac.openMELs.length}  Overdue: ${ac.openMELs.filter((m) => m.isOverdue).length}`,
      `CAME type: Routine Task Finding`,
    ]));
  }

  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ═══════════════════════════════════════════
  // SECTION 6 — COMPONENT RELIABILITY (MTBUR)
  // ═══════════════════════════════════════════

  children.push(sectionHeading("SECTION 6 — COMPONENT RELIABILITY (MTBUR)", "CAME §5.3.4"));

  children.push(para(
    "MTBUR (Mean Time Between Unscheduled Removals) is computed from defect records. " +
    "Components with multiple events in scope are analysed. " +
    "Alert Level = 0.8 × fleet mean MTBUR (flag if below).",
    { italic: true }
  ));

  // Per-aircraft MTBUR from defect records grouped by ATA
  children.push(subHeading("6.1 Fleet MTBUR Summary — CAME §5.3.4"));
  const mtburRows = [];
  for (const [ata, name] of Object.entries(ATA_NAMES)) {
    let fleetRemovals = 0;
    let fleetFH = 0;
    const perAC = acData.map((ac) => {
      const removals = ac.totalFreq[ata] || 0;
      fleetRemovals += removals;
      fleetFH += ac.totalFH;
      const m = mtbur(ac.totalFH, removals);
      return { removals, mtbur: m };
    });
    if (fleetRemovals === 0) continue;

    const fleetMTBUR = mtbur(fleetFH, fleetRemovals);
    const fleetAL    = fleetMTBUR ? +(fleetMTBUR * 0.8).toFixed(1) : null;
    const anyBelow   = perAC.some((a) => a.mtbur != null && fleetAL != null && a.mtbur < fleetAL);

    const row = [ata, name, fleetRemovals, fmt(fleetMTBUR, 1), fmt(fleetAL, 1)];
    row.push(statusCell(anyBelow ? "BELOW AL" : "OK", anyBelow ? "AMBER" : "GREEN"));
    mtburRows.push(row);
  }

  if (mtburRows.length > 0) {
    children.push(simpleTable(
      ["ATA", "Component", "Fleet Removals", "Fleet MTBUR (hrs)", "Fleet AL (hrs)", "Status"],
      mtburRows
    ));
  } else {
    children.push(para("Insufficient data for MTBUR computation. Verify defect records contain component removal events.", { italic: true }));
  }

  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ═══════════════════════════════════════════
  // SECTION 7 — OIL CONSUMPTION
  // ═══════════════════════════════════════════

  children.push(sectionHeading("SECTION 7 — OIL CONSUMPTION", "CAME §5.3.4.2"));

  // Fleet oil summary
  children.push(subHeading("7.1 Fleet Oil Consumption Summary"));
  const oilSummaryRows = acData.map((ac) => {
    const limit = parseFloat(ac.engineProfile.oilLimit) || null;
    const e1rates = ac.fuelMonths.map((m) => m.eng1OilRateLHr).filter((v) => v != null);
    const e2rates = ac.fuelMonths.map((m) => m.eng2OilRateLHr).filter((v) => v != null);
    const e1avg = e1rates.length ? +(e1rates.reduce((a, b) => a + b, 0) / e1rates.length).toFixed(3) : null;
    const e2avg = e2rates.length ? +(e2rates.reduce((a, b) => a + b, 0) / e2rates.length).toFixed(3) : null;
    const e1pct = e1avg && limit ? +((e1avg / limit) * 100).toFixed(1) : null;
    const e2pct = e2avg && limit ? +((e2avg / limit) * 100).toFixed(1) : null;
    const e1ok = e1avg == null || limit == null || e1avg <= limit;
    const e2ok = e2avg == null || limit == null || e2avg <= limit;

    return [
      ac.reg,
      fmt(e1avg, 3), fmt(limit, 3),
      statusCell(e1pct != null ? `${e1pct}%` : "N/A", e1ok ? "GREEN" : "RED"),
      fmt(e2avg, 3), fmt(limit, 3),
      statusCell(e2pct != null ? `${e2pct}%` : "N/A", e2ok ? "GREEN" : "RED"),
      statusCell(
        (!e1avg && !e2avg) ? "No data" : (e1ok && e2ok) ? "NORMAL" : "EXCEEDS LIMIT",
        (!e1avg && !e2avg) ? "AMBER" : (e1ok && e2ok) ? "GREEN" : "RED"
      ),
    ];
  });
  children.push(simpleTable(
    ["Reg", "Eng1 Rate (L/hr)", "Eng1 Limit", "Eng1 %",
     "Eng2 Rate (L/hr)", "Eng2 Limit", "Eng2 %", "Fleet Status"],
    oilSummaryRows
  ));

  // Per-aircraft detail
  for (const ac of acData) {
    children.push(subHeading(`7.2 ${ac.reg} — Oil Consumption Detail`, C.navy));
    const limit = parseFloat(ac.engineProfile.oilLimit) || null;

    if (ac.fuelMonths.length === 0) {
      children.push(para("No fuel/oil data available for this aircraft in scope period.", { italic: true }));
    } else {
      const oilRows = scopeKeys.map((k) => {
        const m = ac.fuelMonths.find((fm) => fm.ymKey === k);
        const label = period.monthsInScope[scopeKeys.indexOf(k)];
        if (!m) return [label, "Not recorded", "Not recorded", "Not recorded", "Not recorded", "Not recorded", "Not recorded"];
        const e1ok = m.eng1OilRateLHr == null || limit == null || m.eng1OilRateLHr <= limit;
        const e2ok = m.eng2OilRateLHr == null || limit == null || m.eng2OilRateLHr <= limit;
        return [
          label,
          fmt(m.eng1OilUpliftL, 2), fmt(m.eng1OilConsL, 2),
          statusCell(fmt(m.eng1OilRateLHr, 3), e1ok ? "GREEN" : "RED"),
          fmt(m.eng2OilUpliftL, 2), fmt(m.eng2OilConsL, 2),
          statusCell(fmt(m.eng2OilRateLHr, 3), e2ok ? "GREEN" : "RED"),
        ];
      });
      children.push(simpleTable(
        ["Month", "Eng1 Uplift (L)", "Eng1 Cons (L)", "Eng1 Rate (L/hr)",
         "Eng2 Uplift (L)", "Eng2 Cons (L)", "Eng2 Rate (L/hr)"],
        oilRows
      ));
    }

    children.push(citationBlock([
      `Aircraft: ${ac.reg}`,
      `File: ${ac.files?.fuelOil || "N/A"}`,
      `Oil Limit: ${ac.engineProfile.oilLimit}`,
      `CAME type: Operational Data`,
      `Note: "Not recorded" ≠ compliance failure`,
    ]));
  }

  // Oil chart
  if (config.enableCharts) {
    const oilLabels = period.monthsInScope;
    const oilDatasets = [];
    for (const [i, ac] of acData.entries()) {
      const limit = parseFloat(ac.engineProfile.oilLimit);
      oilDatasets.push({
        label: `${ac.reg} Eng1`,
        data: scopeKeys.map((k) => {
          const m = ac.fuelMonths.find((fm) => fm.ymKey === k);
          return m?.eng1OilRateLHr ?? null;
        }),
        backgroundColor: `#${acColour(i)}`,
        borderColor: `#${acColour(i)}`,
        type: "bar",
      });
      if (!isNaN(limit) && limit > 0) {
        oilDatasets.push({
          label: `${ac.reg} Limit (${limit})`,
          data: oilLabels.map(() => limit),
          borderColor: "#C00000",
          borderDash: [5, 5],
          borderWidth: 2,
          type: "line",
          fill: false,
          pointRadius: 0,
        });
      }
    }
    const imgOil = await fetchChart({
      ...config,
      chart: {
        type: "bar",
        data: { labels: oilLabels, datasets: oilDatasets },
        options: {
          plugins: { title: { display: true, text: "Chart 7.1 — Engine Oil Consumption per Aircraft [CAME §5.3.4.2]" } },
          scales: { y: { beginAtZero: true, title: { display: true, text: "L/hr" } } },
        },
      },
    });
    children.push(chartImage(imgOil));
  }

  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ═══════════════════════════════════════════
  // SECTION 8 — FUEL PERFORMANCE
  // ═══════════════════════════════════════════

  children.push(sectionHeading("SECTION 8 — FUEL PERFORMANCE", "CAME §5.3.4.1"));

  // Fleet fuel summary
  children.push(subHeading("8.1 Fleet Fuel Summary"));
  const fuelSummaryRows = acData.map((ac) => {
    const rates = ac.fuelMonths.map((m) => m.fuelBurnRateKgHr).filter((v) => v != null);
    const avg = rates.length ? +(rates.reduce((a, b) => a + b, 0) / rates.length).toFixed(1) : null;
    const months = scopeKeys.length;
    const withData = ac.fuelMonths.filter((m) => m.fuelBurnRateKgHr != null).length;
    return [
      ac.reg,
      avg != null ? `${fmt(avg, 1)} kg/hr` : "N/A",
      "See aircraft detail",
      statusCell(avg == null ? "No data" : "Nominal", avg == null ? "AMBER" : "GREEN"),
      withData,
      months - withData,
    ];
  });
  children.push(simpleTable(
    ["Reg", "Avg Burn Rate", "Normal Range", "Status", "Months with Data", "Months Not Recorded"],
    fuelSummaryRows
  ));

  // Per-aircraft fuel tables
  for (const ac of acData) {
    children.push(subHeading(`8.2 ${ac.reg} — Fuel Consumption Detail`, C.navy));
    if (ac.fuelMonths.length === 0) {
      children.push(para("No fuel data available for this aircraft in scope period.", { italic: true }));
    } else {
      const fuelRows = scopeKeys.map((k) => {
        const m = ac.fuelMonths.find((fm) => fm.ymKey === k);
        const label = period.monthsInScope[scopeKeys.indexOf(k)];
        if (!m) return [label, "Not recorded", "Not recorded", "Not recorded", "Not recorded"];
        return [
          label,
          fmt(m.fh, 1),
          fmt(m.fuelUpliftKg, 0),
          fmt(m.fuelBurnKg, 0),
          fmt(m.fuelBurnRateKgHr, 1),
        ];
      });
      children.push(simpleTable(
        ["Month", "FH", "Fuel Uplift (kg)", "Fuel Burn (kg)", "Burn Rate (kg/hr)"],
        fuelRows
      ));
    }
    children.push(citationBlock([
      `Aircraft: ${ac.reg}`,
      `File: ${ac.files?.fuelOil || "N/A"}`,
      `CAME type: Operational Data`,
      `Note: "Not recorded" ≠ compliance failure`,
    ]));
  }

  // Fuel chart
  if (config.enableCharts) {
    const fuelLabels = period.monthsInScope;
    const fuelImg = await fetchChart({
      ...config,
      chart: {
        type: "bar",
        data: {
          labels: fuelLabels,
          datasets: acData.map((ac, i) => ({
            label: ac.reg,
            data: scopeKeys.map((k) => {
              const m = ac.fuelMonths.find((fm) => fm.ymKey === k);
              return m?.fuelBurnRateKgHr ?? null;
            }),
            backgroundColor: `#${acColour(i)}`,
          })),
        },
        options: {
          plugins: { title: { display: true, text: "Chart 8.1 — Monthly Fuel Burn Rate per Aircraft [CAME §5.3.4.1]" } },
          scales: { y: { beginAtZero: false, title: { display: true, text: "kg/hr" } } },
        },
      },
    });
    children.push(chartImage(fuelImg));
  }

  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ═══════════════════════════════════════════
  // SECTION 9 — REGULATORY COMPLIANCE
  // ═══════════════════════════════════════════

  children.push(sectionHeading("SECTION 9 — FLEET REGULATORY COMPLIANCE", "CAME §5.0"));

  const complianceItems = [
    { req: "Monthly Report to NCAA",     ref: "§5.3.3" },
    { req: "Alert Levels reviewed",      ref: "§5.2.9" },
    { req: "RCB Quarterly Meeting",      ref: "§5.0.6" },
    { req: "TRB Weekly Review",          ref: "§5.0.7" },
    { req: "Alert Notices issued",       ref: "§5.2.11" },
    { req: "PIREP/MAREP submitted",      ref: "§5.2.2" },
    { req: "MTBUR report current",       ref: "§5.3.4" },
    { req: "Engine Performance Report",  ref: "§5.3.4.3" },
    { req: "AMP reviewed (if AL breach)",ref: "§5.4.1" },
  ];

  const compRows = complianceItems.map((item) => {
    const note = item.req === "AMP reviewed (if AL breach)"
      ? (allRedFindings.length > 0 ? "Required — AL exceedance detected" : "Not required this period")
      : "";
    const perAC = acData.map((ac) => {
      // Auto-populate based on available data
      if (item.req === "Alert Notices issued") {
        return statusCell(
          ac.hasALBreach ? "REQUIRED" : "N/A",
          ac.hasALBreach ? "RED" : "GREEN"
        );
      }
      if (item.req === "PIREP/MAREP submitted") {
        return statusCell(
          ac.scopePirep.length > 0 ? "Submitted" : "N/A",
          "AMBER"
        );
      }
      return statusCell("Verify", "AMBER");
    });
    return [item.req, item.ref, ...perAC, note];
  });

  children.push(simpleTable(
    ["Requirement", "CAME Ref", ...acData.map((a) => a.reg), "Notes"],
    compRows
  ));

  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ═══════════════════════════════════════════
  // SECTION 10 — RCB / TRB ESCALATION
  // ═══════════════════════════════════════════

  children.push(sectionHeading("SECTION 10 — RCB / TRB ESCALATION STATUS", "CAME §5.0.6 / §5.0.7"));

  children.push(para(
    "All Alert Level exceedances from Section 4 are listed below. " +
    "The RCB reviews findings across the whole fleet.",
    { italic: true }
  ));

  const escalationRows = [];
  for (const f of allRedFindings) {
    escalationRows.push([
      f.reg,
      `AN-${f.reg}-${f.ata}-${new Date().getFullYear()}`,
      f.ata,
      `${ataName(f.ata)} — AL exceedance`,
      "HIGH",
      statusCell("YES", "RED"),
      "TBD",
      statusCell("YES", "AMBER"),
      "Next Quarter",
      statusCell("OPEN", "RED"),
    ]);
  }

  if (escalationRows.length > 0) {
    children.push(simpleTable(
      ["Aircraft", "Finding Ref", "ATA", "Description", "Severity",
       "TRB?", "TRB Date", "RCB Req?", "Target RCB", "Status"],
      escalationRows
    ));
  } else {
    children.push(para("No Alert Level exceedances this period. No RCB escalation required.", { italic: true }));
  }

  children.push(para("Next RCB Meeting Date: [To be confirmed by CAM]", { bold: true }));
  children.push(para("RCB Chairman: [Name]", { bold: true }));

  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ═══════════════════════════════════════════
  // SECTION 11 — RECOMMENDATIONS & ACTION PLAN
  // ═══════════════════════════════════════════

  children.push(sectionHeading("SECTION 11 — FLEET RECOMMENDATIONS & ACTION PLAN", "CAME §5.4"));

  const recommendations = [];

  // Auto-generate from findings
  for (const f of allRedFindings) {
    recommendations.push({
      priority: "HIGH",
      aircraft: f.reg,
      text: `Issue formal Alert Notice for ATA ${f.ata} (${ataName(f.ata)}) — rate exceeds Alert Level. Investigate root cause and submit findings to RCB.`,
      came: "§5.2.11",
      owner: "CAMO Reliability Officer",
      target: fmtDate(addWorkingDays(new Date(), 10)),
      status: "OPEN",
    });
  }

  for (const ac of acData) {
    if (ac.openMELs.some((m) => m.isOverdue)) {
      recommendations.push({
        priority: "HIGH",
        aircraft: ac.reg,
        text: `Review and action overdue MEL items for ${ac.reg}. Ensure deferred defects are cleared within approved categories.`,
        came: "§5.2.5",
        owner: "CAMO / AMO",
        target: fmtDate(addWorkingDays(new Date(), 5)),
        status: "OPEN",
      });
    }
    if (!ac.fuelMonths.length) {
      recommendations.push({
        priority: "MEDIUM",
        aircraft: ac.reg,
        text: `Obtain missing Fuel & Oil log data for ${ac.reg}. Ensure oil consumption monitoring is maintained per CAME.`,
        came: "§5.3.4.2",
        owner: "CAMO Reliability Officer",
        target: "Next reporting period",
        status: "OPEN",
      });
    }
  }

  // Standard recommendations (always included)
  recommendations.push(
    {
      priority: "MEDIUM",
      aircraft: "All Fleet",
      text: "Conduct quarterly RCB meeting to review fleet reliability trends, AL findings, and component MTBUR performance.",
      came: "§5.0.6",
      owner: "Accountable Manager / CAM",
      target: "Quarterly",
      status: "SCHEDULED",
    },
    {
      priority: "MEDIUM",
      aircraft: "All Fleet",
      text: "Maintain weekly TRB review of defect and MEL status across all aircraft. Flag recurring ATA trends for escalation.",
      came: "§5.0.7",
      owner: "Technical Review Board",
      target: "Weekly",
      status: "ONGOING",
    },
    {
      priority: "LOW",
      aircraft: "All Fleet",
      text: "Review data recording quality across all aircraft files. Ensure all months in scope have complete Fuel & Oil, TLP, and defect data.",
      came: "§5.3.3",
      owner: "CAMO Reliability Officer",
      target: "Next month",
      status: "OPEN",
    }
  );

  // Ensure minimum 5
  while (recommendations.length < 5) {
    recommendations.push({
      priority: "LOW",
      aircraft: "All Fleet",
      text: "Review AMP task intervals for top ATA chapters to confirm adequacy of maintenance programme.",
      came: "§5.4.1",
      owner: "CAMO Reliability Officer",
      target: "Next quarter",
      status: "OPEN",
    });
  }

  const recHeaders = ["Priority", "Aircraft", "Recommendation", "CAME Basis", "Owner", "Target Date", "Status"];
  const recRows = recommendations.map((r) => [
    statusCell(r.priority, r.priority === "HIGH" ? "RED" : r.priority === "MEDIUM" ? "AMBER" : "GREEN"),
    r.aircraft,
    r.text,
    r.came,
    r.owner,
    r.target,
    statusCell(r.status, r.status === "OPEN" ? "AMBER" : r.status === "CLOSED" ? "GREEN" : "AMBER"),
  ]);
  children.push(simpleTable(recHeaders, recRows));

  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ═══════════════════════════════════════════
  // SECTION 12 — DATA SOURCES & LINEAGE
  // ═══════════════════════════════════════════

  children.push(sectionHeading("SECTION 12 — DATA SOURCES & LINEAGE", "CAME §5.3"));

  for (const ac of acData) {
    children.push(subHeading(`12.${acData.indexOf(ac) + 1} ${ac.reg} — Data Lineage`, C.navy));

    const fileTypes = [
      { key: "tlp",              type: "TLP",                      data: ac.tlpData,  records: ac.allMonths.length, contrib: "Sections 3, 6" },
      { key: "defectRecords",    type: "Defect Records",            data: ac.defects,  records: ac.defects.length,  contrib: "Sections 4, 6" },
      { key: "deferredDefects",  type: "MEL / Deferred Defect",     data: ac.mel,      records: ac.mel.length,      contrib: "Sections 5"    },
      { key: "fuelOil",          type: "Fuel & Oil Log",            data: ac.fuelOil,  records: ac.fuelMonths.length, contrib: "Sections 7, 8" },
      { key: "delayCancellation",type: "Delay & Cancellation",      data: ac.delays,   records: ac.delays.length,   contrib: "Section 3"     },
      { key: "pirepMarep",       type: "PIREP/MAREP Register",      data: ac.pirep,    records: ac.pirep.length,    contrib: "Sections 3, 9" },
    ];

    const lineageRows = fileTypes.map((f) => {
      const hasFile = !!ac.files[f.key];
      const hasData = f.records > 0;
      const quality = !hasFile ? "File not provided"
        : !hasData ? "File loaded — no records extracted"
        : "OK";
      return [
        ac.files[f.key] || "Not provided",
        f.type,
        hasData ? `${f.records} records` : "0",
        f.contrib,
        statusCell(quality, !hasFile ? "RED" : !hasData ? "AMBER" : "GREEN"),
      ];
    });

    children.push(simpleTable(
      ["File", "Type", "Records", "Contributes To", "Data Quality"],
      lineageRows
    ));

    // Data gaps for this aircraft
    if (ac.dataGaps.length > 0) {
      children.push(para(`Known Data Gaps for ${ac.reg}:`, { bold: true }));
      for (const gap of ac.dataGaps) {
        children.push(para(`  • ${gap.item} — ${gap.month}: ${gap.reason}`, { italic: true }));
      }
    }
  }

  // Fleet data gaps summary
  children.push(subHeading("12.X Fleet Data Gaps Summary"));
  const gapRows = [];
  for (const ac of acData) {
    for (const gap of ac.dataGaps) {
      gapRows.push([ac.reg, gap.item, gap.reason || "Not recorded", "Operational monitoring", "Update next period"]);
    }
    // Auto-detect missing files
    for (const [key, label] of [
      ["delayCancellation", "Delay & Cancellation Log"],
      ["pirepMarep", "PIREP/MAREP Register"],
    ]) {
      if (!ac.files[key]) {
        gapRows.push([ac.reg, label, "File not provided", "Section 3 reliability params", "Provide in next report"]);
      }
    }
  }

  if (gapRows.length > 0) {
    children.push(simpleTable(
      ["Aircraft", "Missing Item", "Reason", "Required By", "Action"],
      gapRows
    ));
  } else {
    children.push(para("No data gaps identified.", { italic: true }));
  }

  children.push(new Paragraph({ children: [new PageBreak()] }));

  // ═══════════════════════════════════════════
  // SECTION 13 — SIGN-OFF
  // ═══════════════════════════════════════════

  children.push(sectionHeading("SECTION 13 — SIGN-OFF", "CAME §5.0.2"));

  children.push(para(
    "This report covers the entire ATTAIR fleet for the period stated on the cover page. " +
    "One sign-off block applies to all aircraft.",
    { italic: true }
  ));

  children.push(
    simpleTable(
      ["Role", "Name", "Signature", "Date"],
      [
        ["Reliability Officer", preparedBy, "___________________", "_____ / _____ / _____"],
        ["Technical Review / CAM", "[Name]", "___________________", "_____ / _____ / _____"],
        ["Quality Assurance", "[Name]", "___________________", "_____ / _____ / _____"],
        ["Accountable Manager", "[Name]", "___________________", "_____ / _____ / _____"],
      ]
    )
  );

  children.push(para(`\nReport Reference: ${reportRef}`, { bold: true }));
  children.push(para(`CAME Reference: ${cameRef}`, { italic: true }));
  children.push(para(`Generated: ${new Date().toLocaleDateString("en-GB", { day: "2-digit", month: "long", year: "numeric" })}`, { italic: true, color: "888888" }));

  // ─── ASSEMBLE DOCUMENT ───────────────────────────────────────────────────

  const doc = new Document({
    creator:     preparedBy,
    title:       `ATTAIR Fleet Reliability Report — ${period.label}`,
    description: `${reportRef} | ${cameRef}`,
    styles: {
      default: {
        document: {
          run: { font: "Calibri", size: 18 },
        },
      },
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top:    convertInchesToTwip(0.8),
              bottom: convertInchesToTwip(0.8),
              left:   convertInchesToTwip(0.8),
              right:  convertInchesToTwip(0.8),
            },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: `ATTAIR Fleet Reliability Report | ${period.label}`, size: 16, color: C.navy }),
                  new TextRun({ text: `    ${reportRef}`, size: 14, color: "888888" }),
                ],
                border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: C.gold } },
              }),
            ],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: `${reportRef}  |  Page `, size: 16, color: "888888" }),
                  new SimpleField("PAGE"),
                  new TextRun({ text: `  |  ATTAIR CAMO  |  ${cameRef}  |  CONFIDENTIAL`, size: 16, color: "888888" }),
                ],
                border: { top: { style: BorderStyle.SINGLE, size: 6, color: C.gold } },
                alignment: AlignmentType.CENTER,
              }),
            ],
          }),
        },
        children,
      },
    ],
  });

  // ─── WRITE FILE ──────────────────────────────────────────────────────────

  const year  = new Date().getFullYear();
  const month = String(new Date().getMonth() + 1).padStart(2, "0");
  const outFile = path.join(OUT_DIR,
    `ATTAIR-CAMO-FLEET-REL-${year}-${month}-${Date.now()}.docx`);

  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(outFile, buf);

  const sizeMB = (buf.length / 1048576).toFixed(2);
  console.log(`\n✅  Report saved: ${outFile}  (${sizeMB} MB)`);

  // ─── POST-GENERATION COMPLIANCE CHECKLIST ────────────────────────────────

  console.log("\n╔══════════════════════════════════════════════════════════╗");
  console.log("║  POST-GENERATION COMPLIANCE CHECKLIST                   ║");
  console.log("╚══════════════════════════════════════════════════════════╝");

  for (const ac of acData) {
    console.log(`\n  [${ac.reg}]`);
    const items = [
      ["TLP / Utilisation data",    ac.allMonths.length > 0],
      ["Defect records",            ac.defects.length > 0],
      ["MEL / Deferred defects",    ac.mel.length > 0],
      ["Fuel & Oil log",            ac.fuelMonths.length > 0],
      ["Delay / Cancel log",        ac.delays.length > 0],
      ["PIREP/MAREP register",      ac.pirep.length > 0],
      ["Alert Level computation",   Object.keys(ac.alertLevels).length > 0],
      ["Engine profile detected",   ac.engineProfile.eng1SN !== "Not detected"],
    ];
    for (const [label, ok] of items) {
      console.log(`    ${ok ? "✓" : "⚠ "} ${label}${ok ? "" : " — missing or empty"}`);
    }
    if (ac.hasALBreach) {
      console.log(`    ⚠  ALERT NOTICE REQUIRED — AL exceedance detected`);
    }
  }

  console.log("\n  Fleet-level:");
  console.log(`    ✓  Sections 1–13 generated`);
  console.log(`    ${allRedFindings.length > 0 ? "⚠  " : "✓  "}${allRedFindings.length} Alert Notice(s) required`);
  console.log(`    ${config.enableCharts ? "✓  Charts enabled (QuickChart.io)" : "⚠  Charts disabled"}`);
  console.log(`    ✓  Document reference: ${reportRef}`);
  console.log(`    ✓  CAME ref: ${cameRef}`);

  if (allRedFindings.length > 0) {
    console.log("\n  ATA chapters with multi-aircraft trend:");
    const multiAC = Object.keys(fleetATATotals).filter((ata) =>
      acData.filter((a) => (a.totalFreq[ata] || 0) > 0).length > 1
    );
    for (const ata of multiAC) {
      console.log(`    ⚠  ATA ${ata} (${ataName(ata)}) — appears in multiple aircraft`);
    }
  }

  console.log(`\n  Output: ${outFile}\n`);
}

main().catch((err) => {
  console.error("\n✗  Fatal error:", err.message);
  console.error(err.stack);
  process.exit(1);
});
