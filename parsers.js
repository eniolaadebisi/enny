/**
 * parsers.js — File parsers for ATTAIR Fleet Reliability Report
 * ==============================================================
 * Reads Excel files produced by airline operations staff and extracts
 * structured data for each aircraft.  All functions are async and
 * return plain objects / arrays that the main script consumes.
 *
 * Supported formats:
 *   TLP            — Technical Log Page workbook (one sheet per month)
 *   Defect Records — Defect register workbook
 *   MEL / DD Sheet — Deferred Defect / MEL workbook
 *   Fuel & Oil Log — Combined fuel and oil consumption workbook
 *   Delay / Cancel — Delay & Cancellation register
 *   PIREP / MAREP  — Pilot / Maintenance report register
 */

"use strict";

const ExcelJS  = require("exceljs");
const fs       = require("fs");
const path     = require("path");

// ─── helpers ────────────────────────────────────────────────────────────────

/** Return trimmed string or null */
function str(v) {
  if (v == null) return null;
  const s = String(v).trim();
  return s === "" ? null : s;
}

/** Return numeric value or null */
function num(v) {
  if (v == null) return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}

/** Return Date or null */
function dat(v) {
  if (v == null) return null;
  if (v instanceof Date) return v;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

/** YYYY-MM string from Date */
function ym(d) {
  if (!d) return null;
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

/** "Month YYYY" → "YYYY-MM" for month matching */
function monthLabel(label) {
  const months = {
    january:"01", february:"02", march:"03", april:"04",
    may:"05", june:"06", july:"07", august:"08",
    september:"09", october:"10", november:"11", december:"12",
  };
  const m = label.trim().match(/^(\w+)\s+(\d{4})$/);
  if (!m) return label;
  const mo = months[m[1].toLowerCase()];
  return mo ? `${m[2]}-${mo}` : label;
}

/** Resolve file path; return null and log warning if missing */
function resolvePath(filePath, reg, type) {
  if (!filePath) return null;
  const resolved = path.resolve(filePath);
  if (!fs.existsSync(resolved)) {
    console.warn(`  ⚠  [${reg}] ${type} file not found: ${filePath}`);
    return null;
  }
  return resolved;
}

/** Load workbook; return null on failure */
async function loadWorkbook(filePath) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(filePath);
  return wb;
}

/** Collect all non-empty rows from a sheet as arrays of cell values */
function sheetRows(sheet) {
  const rows = [];
  sheet.eachRow({ includeEmpty: false }, (row) => {
    rows.push(row.values.slice(1)); // index 0 is always undefined in ExcelJS
  });
  return rows;
}

/** Find a row index (0-based) whose first cell matches a pattern */
function findRowIndex(rows, pattern) {
  for (let i = 0; i < rows.length; i++) {
    const v = str(rows[i][0]);
    if (v && pattern.test(v)) return i;
  }
  return -1;
}

// ─── TLP PARSER ─────────────────────────────────────────────────────────────

/**
 * Parse a TLP workbook.
 *
 * Returns:
 * {
 *   engineProfile: { eng1SN, eng2SN, oilLimit, apuSN },
 *   months: [
 *     {
 *       label,       // e.g. "January 2026"
 *       ymKey,       // e.g. "2026-01"
 *       fh, cyc,     // airframe FH / cycles for the month
 *       eng1FH, eng1CYC,
 *       eng2FH, eng2CYC,
 *       apuHrs, apuStarts,
 *       sectors,     // array of { date, from, to, blockOff, blockOn, sectorFH }
 *       operatingDays,
 *       nilOpDays,
 *       calDays,
 *       openingTSN, openingCSN,
 *       closingTSN,  closingCSN,
 *       isGroundPeriod,
 *     }
 *   ],
 *   rawSheetNames: [],
 * }
 */
async function parseTLP(filePath, registration, monthsInScope) {
  const resolved = resolvePath(filePath, registration, "TLP");
  if (!resolved) return null;

  const wb = await loadWorkbook(resolved);
  const scopeKeys = monthsInScope.map(monthLabel);
  const engineProfile = { eng1SN: null, eng2SN: null, oilLimit: null, apuSN: null };
  const months = [];

  for (const sheet of wb.worksheets) {
    const rows = sheetRows(sheet);
    if (rows.length < 3) continue;

    // ── Auto-detect engine profile from header rows ────────────────
    if (!engineProfile.eng1SN) {
      for (let i = 0; i < Math.min(10, rows.length); i++) {
        for (let j = 0; j < rows[i].length; j++) {
          const v = str(rows[i][j]);
          if (!v) continue;
          if (/eng(?:ine)?\s*1\s*s[/.]?n/i.test(v) && rows[i][j + 1]) {
            engineProfile.eng1SN = str(rows[i][j + 1]);
          }
          if (/eng(?:ine)?\s*2\s*s[/.]?n/i.test(v) && rows[i][j + 1]) {
            engineProfile.eng2SN = str(rows[i][j + 1]);
          }
          if (/apu\s*s[/.]?n/i.test(v) && rows[i][j + 1]) {
            engineProfile.apuSN = str(rows[i][j + 1]);
          }
          if (/oil\s*limit/i.test(v) && rows[i][j + 1]) {
            engineProfile.oilLimit = str(rows[i][j + 1]);
          }
        }
      }
    }

    // ── Identify sheet month ───────────────────────────────────────
    // Sheet name or header cell may contain month label
    let sheetMonthLabel = null;
    let sheetYM = null;

    const sheetName = sheet.name.trim();
    // Try sheet name first
    const snMatch = sheetName.match(/(\w+)\s+(\d{4})/);
    if (snMatch) {
      sheetMonthLabel = `${snMatch[1]} ${snMatch[2]}`;
      sheetYM = monthLabel(sheetMonthLabel);
    }
    // Try first row cell
    if (!sheetYM) {
      for (let j = 0; j < (rows[0] || []).length; j++) {
        const v = str(rows[0][j]);
        if (!v) continue;
        const rm = v.match(/(\w+)\s+(\d{4})/);
        if (rm) {
          sheetMonthLabel = `${rm[1]} ${rm[2]}`;
          sheetYM = monthLabel(sheetMonthLabel);
          break;
        }
      }
    }
    if (!sheetYM) continue; // Cannot identify month — skip

    // ── Locate header row (contains FROM / TO column labels) ───────
    let headerRowIdx = -1;
    for (let i = 0; i < Math.min(20, rows.length); i++) {
      const rowStr = rows[i].map((v) => str(v) || "").join("|").toLowerCase();
      if (rowStr.includes("from") && rowStr.includes("to")) {
        headerRowIdx = i;
        break;
      }
    }
    if (headerRowIdx === -1) continue;

    const headerRow = rows[headerRowIdx].map((v) => str(v) || "");

    // Map column indices
    const col = {
      date:       headerRow.findIndex((h) => /date/i.test(h)),
      from:       headerRow.findIndex((h) => /^from$/i.test(h)),
      to:         headerRow.findIndex((h) => /^to$/i.test(h)),
      blockOff:   headerRow.findIndex((h) => /block.*off|off.*block/i.test(h)),
      blockOn:    headerRow.findIndex((h) => /block.*on|on.*block/i.test(h)),
      secFH:      headerRow.findIndex((h) => /sector.*fh|flight.*hrs/i.test(h)),
      // Airframe running totals
      afTSN:      headerRow.findIndex((h) => /airframe.*tsn|a\/f.*tsn/i.test(h)),
      afCSN:      headerRow.findIndex((h) => /airframe.*csn|a\/f.*csn/i.test(h)),
      // Engine columns
      e1TSN:      headerRow.findIndex((h) => /eng.*1.*tsn|e1.*tsn/i.test(h)),
      e1CSN:      headerRow.findIndex((h) => /eng.*1.*csn|e1.*csn/i.test(h)),
      e2TSN:      headerRow.findIndex((h) => /eng.*2.*tsn|e2.*tsn/i.test(h)),
      e2CSN:      headerRow.findIndex((h) => /eng.*2.*csn|e2.*csn/i.test(h)),
      // APU
      apuHrs:     headerRow.findIndex((h) => /apu.*hrs|apu.*hours/i.test(h)),
      apuStarts:  headerRow.findIndex((h) => /apu.*start/i.test(h)),
    };

    // ── Locate B/FWD and TOTAL rows ────────────────────────────────
    const bfwdIdx  = findRowIndex(rows, /b\/fwd|b\.fwd|brought.?forward/i);
    const totalIdx = findRowIndex(rows, /^total$/i);

    // ── Extract B/FWD values ───────────────────────────────────────
    const getBFWD = (colIdx) =>
      bfwdIdx >= 0 && colIdx >= 0 ? num(rows[bfwdIdx][colIdx]) : null;
    const getTotal = (colIdx) =>
      totalIdx >= 0 && colIdx >= 0 ? num(rows[totalIdx][colIdx]) : null;

    const openingTSN = getBFWD(col.afTSN);
    const openingCSN = getBFWD(col.afCSN);
    const closingTSN = getTotal(col.afTSN);
    const closingCSN = getTotal(col.afCSN);

    const monthFH  = (closingTSN != null && openingTSN != null) ? +(closingTSN - openingTSN).toFixed(2) : null;
    const monthCYC = (closingCSN != null && openingCSN != null) ? Math.round(closingCSN - openingCSN) : null;

    const eng1FH  = (col.e1TSN >= 0)
      ? safeSubtract(getTotal(col.e1TSN), getBFWD(col.e1TSN))
      : null;
    const eng1CYC = (col.e1CSN >= 0)
      ? safeSubtract(getTotal(col.e1CSN), getBFWD(col.e1CSN))
      : null;
    const eng2FH  = (col.e2TSN >= 0)
      ? safeSubtract(getTotal(col.e2TSN), getBFWD(col.e2TSN))
      : null;
    const eng2CYC = (col.e2CSN >= 0)
      ? safeSubtract(getTotal(col.e2CSN), getBFWD(col.e2CSN))
      : null;

    const apuHrsMonth    = (col.apuHrs >= 0)
      ? safeSubtract(getTotal(col.apuHrs), getBFWD(col.apuHrs))
      : null;
    const apuStartsMonth = (col.apuStarts >= 0)
      ? safeSubtract(getTotal(col.apuStarts), getBFWD(col.apuStarts))
      : null;

    // ── Extract sector rows ────────────────────────────────────────
    const sectors = [];
    const operatingDates = new Set();

    for (let i = headerRowIdx + 1; i < rows.length; i++) {
      if (i === bfwdIdx || i === totalIdx) continue;
      const row = rows[i];
      const fromVal = str(row[col.from]);
      const toVal   = col.to >= 0 ? str(row[col.to]) : null;
      if (!fromVal) continue; // not a sector row

      const sectorDate = col.date >= 0 ? dat(row[col.date]) : null;
      if (sectorDate) operatingDates.add(sectorDate.toDateString());

      sectors.push({
        date:     sectorDate,
        from:     fromVal,
        to:       toVal,
        blockOff: col.blockOff >= 0 ? str(row[col.blockOff]) : null,
        blockOn:  col.blockOn  >= 0 ? str(row[col.blockOn])  : null,
        sectorFH: col.secFH   >= 0 ? num(row[col.secFH])    : null,
      });
    }

    // ── Derive calendar / operating / nil-op days ──────────────────
    const ymParts = sheetYM.split("-");
    const calDays = new Date(
      parseInt(ymParts[0]),
      parseInt(ymParts[1]),
      0
    ).getDate();
    const operatingDays = operatingDates.size;
    const nilOpDays     = calDays - operatingDays;

    months.push({
      label:         sheetMonthLabel,
      ymKey:         sheetYM,
      fh:            monthFH  ?? (sectors.length === 0 ? 0 : null),
      cyc:           monthCYC ?? (sectors.length === 0 ? 0 : null),
      eng1FH,  eng1CYC,
      eng2FH,  eng2CYC,
      apuHrs:        apuHrsMonth,
      apuStarts:     apuStartsMonth,
      sectors,
      operatingDays,
      nilOpDays,
      calDays,
      openingTSN,    openingCSN,
      closingTSN,    closingCSN,
      isGroundPeriod: sectors.length === 0,
      scopeMonth:    scopeKeys.includes(sheetYM),
    });
  }

  return {
    engineProfile,
    months,
    rawSheetNames: wb.worksheets.map((s) => s.name),
  };
}

function safeSubtract(a, b) {
  if (a == null || b == null) return null;
  return +(a - b).toFixed(2);
}

// ─── DEFECT RECORDS PARSER ──────────────────────────────────────────────────

/**
 * Parse defect records workbook.
 *
 * Returns array of defect objects:
 * { date, ata, description, rectification, reference, enteredBy, closedDate, status }
 */
async function parseDefectRecords(filePath, registration) {
  const resolved = resolvePath(filePath, registration, "Defect Records");
  if (!resolved) return [];

  const wb = await loadWorkbook(resolved);
  const defects = [];

  for (const sheet of wb.worksheets) {
    const rows = sheetRows(sheet);
    if (rows.length < 2) continue;

    // Find header row
    let hdrIdx = -1;
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const rowStr = rows[i].map((v) => str(v) || "").join("|").toLowerCase();
      if (rowStr.includes("ata") || rowStr.includes("defect") || rowStr.includes("description")) {
        hdrIdx = i;
        break;
      }
    }
    if (hdrIdx === -1) hdrIdx = 0;

    const hdr = rows[hdrIdx].map((v) => str(v) || "");
    const c = {
      date:          hdr.findIndex((h) => /date/i.test(h) && !/close|rect/i.test(h)),
      ata:           hdr.findIndex((h) => /^ata/i.test(h)),
      description:   hdr.findIndex((h) => /desc|defect/i.test(h)),
      rectification: hdr.findIndex((h) => /rect|action/i.test(h)),
      reference:     hdr.findIndex((h) => /ref(?:erence)?|dd\s*no|item/i.test(h)),
      enteredBy:     hdr.findIndex((h) => /entered|raised|by/i.test(h)),
      closedDate:    hdr.findIndex((h) => /clos/i.test(h)),
      status:        hdr.findIndex((h) => /status/i.test(h)),
    };

    for (let i = hdrIdx + 1; i < rows.length; i++) {
      const row = rows[i];
      const ataRaw = c.ata >= 0 ? str(row[c.ata]) : null;
      const desc   = c.description >= 0 ? str(row[c.description]) : null;
      if (!ataRaw && !desc) continue;

      defects.push({
        date:          c.date >= 0 ? dat(row[c.date]) : null,
        ata:           normaliseATA(ataRaw),
        description:   desc,
        rectification: c.rectification >= 0 ? str(row[c.rectification]) : null,
        reference:     c.reference >= 0 ? str(row[c.reference]) : null,
        enteredBy:     c.enteredBy >= 0 ? str(row[c.enteredBy]) : null,
        closedDate:    c.closedDate >= 0 ? dat(row[c.closedDate]) : null,
        status:        c.status >= 0 ? str(row[c.status]) : null,
        sheet:         sheet.name,
      });
    }
  }

  return defects;
}

// ─── MEL / DEFERRED DEFECT PARSER ───────────────────────────────────────────

/**
 * Returns array of MEL items:
 * { ref, ata, description, category, raisedDate, dueDate, closedDate,
 *   status, daysOpen, isOverdue }
 */
async function parseMEL(filePath, registration) {
  const resolved = resolvePath(filePath, registration, "MEL/DD Sheet");
  if (!resolved) return [];

  const wb = await loadWorkbook(resolved);
  const items = [];

  for (const sheet of wb.worksheets) {
    const rows = sheetRows(sheet);
    if (rows.length < 2) continue;

    let hdrIdx = -1;
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const rowStr = rows[i].map((v) => str(v) || "").join("|").toLowerCase();
      if (rowStr.includes("ata") || rowStr.includes("mel") || rowStr.includes("defer")) {
        hdrIdx = i;
        break;
      }
    }
    if (hdrIdx === -1) hdrIdx = 0;

    const hdr = rows[hdrIdx].map((v) => str(v) || "");
    const c = {
      ref:         hdr.findIndex((h) => /ref(?:erence)?|item|dd\s*no|mel\s*no/i.test(h)),
      ata:         hdr.findIndex((h) => /^ata/i.test(h)),
      description: hdr.findIndex((h) => /desc|defect/i.test(h)),
      category:    hdr.findIndex((h) => /cat(?:egory)?/i.test(h)),
      raisedDate:  hdr.findIndex((h) => /raised|opened|date/i.test(h)),
      dueDate:     hdr.findIndex((h) => /due|expir/i.test(h)),
      closedDate:  hdr.findIndex((h) => /clos/i.test(h)),
      status:      hdr.findIndex((h) => /status/i.test(h)),
    };

    const today = new Date();

    for (let i = hdrIdx + 1; i < rows.length; i++) {
      const row = rows[i];
      const ataRaw = c.ata >= 0 ? str(row[c.ata]) : null;
      const desc   = c.description >= 0 ? str(row[c.description]) : null;
      if (!ataRaw && !desc) continue;

      const raisedDate = c.raisedDate >= 0 ? dat(row[c.raisedDate]) : null;
      const closedDate = c.closedDate >= 0 ? dat(row[c.closedDate]) : null;
      const dueDate    = c.dueDate >= 0 ? dat(row[c.dueDate]) : null;
      const statusRaw  = c.status >= 0 ? str(row[c.status]) : null;
      const isOpen     = !closedDate && !/closed|cleared|rectified/i.test(statusRaw || "");
      const daysOpen   = raisedDate
        ? Math.round((today - raisedDate) / 86400000)
        : null;
      const isOverdue  = isOpen && dueDate ? today > dueDate : false;

      items.push({
        ref:         c.ref >= 0 ? str(row[c.ref]) : null,
        ata:         normaliseATA(ataRaw),
        description: desc,
        category:    c.category >= 0 ? str(row[c.category]) : null,
        raisedDate,
        dueDate,
        closedDate,
        status:      statusRaw ?? (isOpen ? "OPEN" : "CLOSED"),
        daysOpen,
        isOpen,
        isOverdue,
        sheet:       sheet.name,
      });
    }
  }

  return items;
}

// ─── FUEL & OIL LOG PARSER ──────────────────────────────────────────────────

/**
 * Returns:
 * {
 *   engineProfile: { eng1SN, eng2SN, oilLimit, apuSN },
 *   months: [
 *     {
 *       label, ymKey,
 *       fuelUpliftKg, fuelBurnKg, fuelBurnRateKgHr,
 *       eng1OilUpliftL, eng1OilConsL, eng1OilRateLHr,
 *       eng2OilUpliftL, eng2OilConsL, eng2OilRateLHr,
 *       apuOilUpliftL,  apuOilConsL,
 *       fh,
 *     }
 *   ]
 * }
 */
async function parseFuelOil(filePath, registration) {
  const resolved = resolvePath(filePath, registration, "Fuel & Oil Log");
  if (!resolved) return null;

  const wb = await loadWorkbook(resolved);
  const engineProfile = { eng1SN: null, eng2SN: null, oilLimit: null, apuSN: null };
  const months = [];

  // Engine profile from header rows (rows 1-5)
  const firstSheet = wb.worksheets[0];
  if (firstSheet) {
    const rows = sheetRows(firstSheet);
    for (let i = 0; i < Math.min(5, rows.length); i++) {
      for (let j = 0; j < rows[i].length; j++) {
        const v = str(rows[i][j]);
        if (!v) continue;
        if (/eng(?:ine)?\s*1\s*s[/.]?n/i.test(v) && rows[i][j + 1])
          engineProfile.eng1SN = str(rows[i][j + 1]);
        if (/eng(?:ine)?\s*2\s*s[/.]?n/i.test(v) && rows[i][j + 1])
          engineProfile.eng2SN = str(rows[i][j + 1]);
        if (/apu\s*s[/.]?n/i.test(v) && rows[i][j + 1])
          engineProfile.apuSN = str(rows[i][j + 1]);
        if (/oil\s*limit/i.test(v) && rows[i][j + 1])
          engineProfile.oilLimit = str(rows[i][j + 1]);
      }
    }
  }

  // One sheet or tab per month OR all months in one sheet
  for (const sheet of wb.worksheets) {
    const rows = sheetRows(sheet);
    if (rows.length < 2) continue;

    // Find header row
    let hdrIdx = -1;
    for (let i = 0; i < Math.min(15, rows.length); i++) {
      const rowStr = rows[i].map((v) => str(v) || "").join("|").toLowerCase();
      if (rowStr.includes("fuel") || rowStr.includes("oil") || rowStr.includes("month")) {
        hdrIdx = i;
        break;
      }
    }
    if (hdrIdx === -1) continue;

    const hdr = rows[hdrIdx].map((v) => str(v) || "");
    const c = {
      month:         hdr.findIndex((h) => /month|period/i.test(h)),
      fuelUplift:    hdr.findIndex((h) => /fuel.*uplift|uplift.*fuel/i.test(h)),
      fuelBurn:      hdr.findIndex((h) => /fuel.*burn|burn.*fuel|consumption.*fuel/i.test(h)),
      fuelBurnRate:  hdr.findIndex((h) => /burn.*rate|rate.*burn|kg.*hr|fuel.*rate/i.test(h)),
      fh:            hdr.findIndex((h) => /\bfh\b|flight.*hr|hours?.*flown/i.test(h)),
      e1Uplift:      hdr.findIndex((h) => /e(?:ng)?1.*oil.*uplift|eng.*1.*uplift/i.test(h)),
      e1Cons:        hdr.findIndex((h) => /e(?:ng)?1.*oil.*cons|e(?:ng)?1.*consum/i.test(h)),
      e1Rate:        hdr.findIndex((h) => /e(?:ng)?1.*oil.*rate|e(?:ng)?1.*l.*hr/i.test(h)),
      e2Uplift:      hdr.findIndex((h) => /e(?:ng)?2.*oil.*uplift|eng.*2.*uplift/i.test(h)),
      e2Cons:        hdr.findIndex((h) => /e(?:ng)?2.*oil.*cons|e(?:ng)?2.*consum/i.test(h)),
      e2Rate:        hdr.findIndex((h) => /e(?:ng)?2.*oil.*rate|e(?:ng)?2.*l.*hr/i.test(h)),
      apuUplift:     hdr.findIndex((h) => /apu.*uplift/i.test(h)),
      apuCons:       hdr.findIndex((h) => /apu.*cons/i.test(h)),
    };

    for (let i = hdrIdx + 1; i < rows.length; i++) {
      const row = rows[i];
      const monthRaw = c.month >= 0 ? str(row[c.month]) : null;
      if (!monthRaw) continue;
      const ymKey = monthLabel(monthRaw);
      if (!/^\d{4}-\d{2}$/.test(ymKey)) continue; // skip non-month rows

      months.push({
        label:           monthRaw,
        ymKey,
        fuelUpliftKg:    c.fuelUplift >= 0 ? num(row[c.fuelUplift]) : null,
        fuelBurnKg:      c.fuelBurn >= 0 ? num(row[c.fuelBurn]) : null,
        fuelBurnRateKgHr:c.fuelBurnRate >= 0 ? num(row[c.fuelBurnRate]) : null,
        fh:              c.fh >= 0 ? num(row[c.fh]) : null,
        eng1OilUpliftL:  c.e1Uplift >= 0 ? num(row[c.e1Uplift]) : null,
        eng1OilConsL:    c.e1Cons >= 0 ? num(row[c.e1Cons]) : null,
        eng1OilRateLHr:  c.e1Rate >= 0 ? num(row[c.e1Rate]) : null,
        eng2OilUpliftL:  c.e2Uplift >= 0 ? num(row[c.e2Uplift]) : null,
        eng2OilConsL:    c.e2Cons >= 0 ? num(row[c.e2Cons]) : null,
        eng2OilRateLHr:  c.e2Rate >= 0 ? num(row[c.e2Rate]) : null,
        apuOilUpliftL:   c.apuUplift >= 0 ? num(row[c.apuUplift]) : null,
        apuOilConsL:     c.apuCons >= 0 ? num(row[c.apuCons]) : null,
        sheet:           sheet.name,
      });
    }
  }

  return { engineProfile, months };
}

// ─── DELAY & CANCELLATION PARSER ────────────────────────────────────────────

/**
 * Returns array:
 * { date, flightNo, from, to, delayMins, reason, ataChapter, type }
 * type: "DELAY" | "CANCELLATION"
 */
async function parseDelayCancellation(filePath, registration) {
  const resolved = resolvePath(filePath, registration, "Delay & Cancellation");
  if (!resolved) return [];

  const wb = await loadWorkbook(resolved);
  const items = [];

  for (const sheet of wb.worksheets) {
    const rows = sheetRows(sheet);
    if (rows.length < 2) continue;

    let hdrIdx = -1;
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const rowStr = rows[i].map((v) => str(v) || "").join("|").toLowerCase();
      if (rowStr.includes("delay") || rowStr.includes("cancel") || rowStr.includes("flight")) {
        hdrIdx = i;
        break;
      }
    }
    if (hdrIdx === -1) hdrIdx = 0;

    const hdr = rows[hdrIdx].map((v) => str(v) || "");
    const c = {
      date:      hdr.findIndex((h) => /date/i.test(h)),
      flightNo:  hdr.findIndex((h) => /flight|flt/i.test(h)),
      from:      hdr.findIndex((h) => /^from|origin|dep/i.test(h)),
      to:        hdr.findIndex((h) => /^to|dest|arr/i.test(h)),
      delayMins: hdr.findIndex((h) => /delay.*min|min.*delay/i.test(h)),
      reason:    hdr.findIndex((h) => /reason|remark/i.test(h)),
      ata:       hdr.findIndex((h) => /^ata/i.test(h)),
      type:      hdr.findIndex((h) => /type|cat(?:egory)?/i.test(h)),
    };

    for (let i = hdrIdx + 1; i < rows.length; i++) {
      const row = rows[i];
      const d = c.date >= 0 ? dat(row[c.date]) : null;
      if (!d) continue;

      const typeRaw = c.type >= 0 ? str(row[c.type]) : null;
      const typeVal = /cancel/i.test(typeRaw || "") ? "CANCELLATION" : "DELAY";

      items.push({
        date:      d,
        flightNo:  c.flightNo >= 0 ? str(row[c.flightNo]) : null,
        from:      c.from >= 0 ? str(row[c.from]) : null,
        to:        c.to >= 0 ? str(row[c.to]) : null,
        delayMins: c.delayMins >= 0 ? num(row[c.delayMins]) : null,
        reason:    c.reason >= 0 ? str(row[c.reason]) : null,
        ata:       normaliseATA(c.ata >= 0 ? str(row[c.ata]) : null),
        type:      typeVal,
        sheet:     sheet.name,
      });
    }
  }

  return items;
}

// ─── PIREP / MAREP PARSER ───────────────────────────────────────────────────

/**
 * Returns array:
 * { date, reportNo, type, ata, description, action, status }
 * type: "PIREP" | "MAREP"
 */
async function parsePirepMarep(filePath, registration) {
  const resolved = resolvePath(filePath, registration, "PIREP/MAREP");
  if (!resolved) return [];

  const wb = await loadWorkbook(resolved);
  const items = [];

  for (const sheet of wb.worksheets) {
    const rows = sheetRows(sheet);
    if (rows.length < 2) continue;

    let hdrIdx = -1;
    for (let i = 0; i < Math.min(10, rows.length); i++) {
      const rowStr = rows[i].map((v) => str(v) || "").join("|").toLowerCase();
      if (rowStr.includes("pirep") || rowStr.includes("marep") || rowStr.includes("report")) {
        hdrIdx = i;
        break;
      }
    }
    if (hdrIdx === -1) hdrIdx = 0;

    const hdr = rows[hdrIdx].map((v) => str(v) || "");
    const c = {
      date:      hdr.findIndex((h) => /date/i.test(h)),
      reportNo:  hdr.findIndex((h) => /ref|no\.?|report.*no/i.test(h)),
      type:      hdr.findIndex((h) => /type/i.test(h)),
      ata:       hdr.findIndex((h) => /^ata/i.test(h)),
      description: hdr.findIndex((h) => /desc|fault|snag/i.test(h)),
      action:    hdr.findIndex((h) => /action|rect/i.test(h)),
      status:    hdr.findIndex((h) => /status/i.test(h)),
    };

    for (let i = hdrIdx + 1; i < rows.length; i++) {
      const row = rows[i];
      const d = c.date >= 0 ? dat(row[c.date]) : null;
      if (!d) continue;

      const typeRaw = c.type >= 0 ? str(row[c.type]) : null;
      const typeVal = /marep/i.test(typeRaw || "") ? "MAREP" : "PIREP";

      items.push({
        date:        d,
        reportNo:    c.reportNo >= 0 ? str(row[c.reportNo]) : null,
        type:        typeVal,
        ata:         normaliseATA(c.ata >= 0 ? str(row[c.ata]) : null),
        description: c.description >= 0 ? str(row[c.description]) : null,
        action:      c.action >= 0 ? str(row[c.action]) : null,
        status:      c.status >= 0 ? str(row[c.status]) : null,
        sheet:       sheet.name,
      });
    }
  }

  return items;
}

// ─── ATA NORMALISATION ──────────────────────────────────────────────────────

const ATA_NAMES = {
  "05":"Time/Limits",
  "21":"Air Conditioning & Pressurization",
  "22":"Auto Flight",
  "23":"Communications",
  "24":"Electrical Power",
  "25":"Equipment & Furnishings",
  "26":"Fire Protection",
  "27":"Flight Controls",
  "28":"Fuel",
  "29":"Hydraulic Power",
  "30":"Ice & Rain Protection",
  "31":"Indicating/Recording",
  "32":"Landing Gear",
  "33":"Lights",
  "34":"Navigation",
  "35":"Oxygen",
  "36":"Pneumatics",
  "38":"Water/Waste",
  "49":"APU",
  "71":"Powerplant",
  "72":"Engine",
  "73":"Engine Fuel & Control",
  "74":"Engine Ignition",
  "75":"Engine Air",
  "76":"Engine Controls",
  "77":"Engine Indicating",
  "78":"Engine Exhaust",
  "79":"Engine Oil",
  "80":"Engine Starting",
};

function normaliseATA(raw) {
  if (!raw) return null;
  const m = raw.match(/(\d{2})/);
  if (!m) return raw;
  const chapter = m[1];
  return ATA_NAMES[chapter] ? chapter : raw;
}

function ataName(chapter) {
  return ATA_NAMES[chapter] || "Unknown";
}

// ─── EXPORTS ────────────────────────────────────────────────────────────────

module.exports = {
  parseTLP,
  parseDefectRecords,
  parseMEL,
  parseFuelOil,
  parseDelayCancellation,
  parsePirepMarep,
  monthLabel,
  normaliseATA,
  ataName,
  ATA_NAMES,
};
