/**
 * POS File Checker — script.js
 *
 * Pure vanilla JavaScript. No frameworks, no build tools.
 * All business logic ported directly from the VBA macros in
 * POSFileChecker_v1.xlsm. Runs entirely in the browser;
 * files are never sent to any server.
 *
 * UPDATED:
 *   - Supports selecting / dropping MULTIPLE files at once.
 *     Every file is parsed and validated independently and the
 *     results are kept in a per-file collection (`allResults`).
 *   - Validation comparisons now use a ±0.01 tolerance: a check
 *     only fails if the absolute difference between the tenant
 *     value and the admin-expected value is GREATER than 0.01.
 *     A difference of exactly 0.01 (or less) is treated as Pass.
 *
 * Sections:
 *   1.  Line item definitions   (mirrors Column A+B of spreadsheet)
 *   2.  File decoder            (UTF-8 → windows-1252 → iso-8859-1 fallback)
 *   3.  File parser             (.001 fixed-width format → structured data)
 *   4.  Calculator              (11 admin-expected values, VBA formulas)
 *   5.  Validator               (orchestrates pipeline, builds result object,
 *                                 ±0.01 tolerance comparison)
 *   6.  Formatters               (currency, date display)
 *   7.  PDF export               (via jsPDF + autoTable from CDN)
 *   8.  Excel export             (via SheetJS from CDN)
 *   9.  UI rendering             (drop zone, summary view, detail view, filters)
 *   10. Event wiring             (drag-drop, file input, multi-file batch, reset)
 */

'use strict';

/* ============================================================
   1. LINE ITEM DEFINITIONS
   Static metadata for all 65 .001 file positions.
   decodeMode: 'currency' → raw int ÷ 100 = PHP amount
               'integer'  → plain integer, no division
               'text'     → raw string
   isValidated: true → Column E has Pass/Failed check
   section: 'vat' (lines 1-34) | 'nonvat' (lines 35-65)
   ============================================================ */
const LINE_ITEMS = [
  // ── Section A: VAT Sales (lines 1–34) ──────────────────────────────────
  { lineItem:  1, definition: 'Tenant Code',                              decodeMode: 'text',     isValidated: false, section: 'vat' },
  { lineItem:  2, definition: 'POS Terminal Number',                      decodeMode: 'text',     isValidated: false, section: 'vat' },
  { lineItem:  3, definition: 'Date (mmddyyyy)',                          decodeMode: 'text',     isValidated: false, section: 'vat' },
  { lineItem:  4, definition: 'Old Accumulated Sales',                    decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem:  5, definition: 'New Accumulated Sales',                    decodeMode: 'currency', isValidated: true,  section: 'vat' },
  { lineItem:  6, definition: 'Total Gross Amount',                       decodeMode: 'currency', isValidated: true,  section: 'vat' },
  { lineItem:  7, definition: 'Total Deductions',                         decodeMode: 'currency', isValidated: true,  section: 'vat' },
  { lineItem:  8, definition: 'Total Promo Sales Amount',                 decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem:  9, definition: 'Total Discount',                           decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 10, definition: 'Total Refund Amount',                      decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 11, definition: 'Total Returned Items Amount',              decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 12, definition: 'Total Other Taxes',                        decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 13, definition: 'Total Service Charge Amount',              decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 14, definition: 'Total Adjustment Discount',                decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 15, definition: 'Total Void Amount',                        decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 16, definition: 'Total Discount Cards',                     decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 17, definition: 'Total Delivery Charges',                   decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 18, definition: 'Total Gift Certificates/Checks Redeemed', decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 19, definition: 'Store Specific Discount 1 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 20, definition: 'Store Specific Discount 2 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 21, definition: 'Store Specific Discount 3 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 22, definition: 'Store Specific Discount 4 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 23, definition: 'Store Specific Discount 5 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 24, definition: 'Total of all Non-Approved Store Discounts',decodeMode: 'currency', isValidated: true,  section: 'vat' },
  { lineItem: 25, definition: 'Store Specific Discount 1 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 26, definition: 'Store Specific Discount 2 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 27, definition: 'Store Specific Discount 3 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 28, definition: 'Store Specific Discount 4 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 29, definition: 'Store Specific Discount 5 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'vat' },
  { lineItem: 30, definition: 'Total VAT/Tax Amount',                     decodeMode: 'currency', isValidated: true,  section: 'vat' },
  { lineItem: 31, definition: 'Total Net Sales Amount',                   decodeMode: 'currency', isValidated: true,  section: 'vat' },
  // Lines 32–34: plain integers — VBA branch `If i >= 31 And i <= 34`
  { lineItem: 32, definition: 'Total Cover Count',                        decodeMode: 'integer',  isValidated: false, section: 'vat' },
  { lineItem: 33, definition: 'Control Number',                           decodeMode: 'integer',  isValidated: false, section: 'vat' },
  { lineItem: 34, definition: 'Total Number of Sales Transactions',       decodeMode: 'integer',  isValidated: false, section: 'vat' },
  // ── Section B: Non-VAT Sales (lines 35–65) ─────────────────────────────
  { lineItem: 35, definition: 'Sales Type',                               decodeMode: 'text',     isValidated: false, section: 'nonvat' },
  { lineItem: 36, definition: 'Amount',                                   decodeMode: 'currency', isValidated: true,  section: 'nonvat' },
  { lineItem: 37, definition: 'Old Accumulated Sales',                    decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 38, definition: 'New Accumulated Sales',                    decodeMode: 'currency', isValidated: true,  section: 'nonvat' },
  { lineItem: 39, definition: 'Total Gross Amount',                       decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 40, definition: 'Total Deductions',                         decodeMode: 'currency', isValidated: true,  section: 'nonvat' },
  { lineItem: 41, definition: 'Total Promo Sales Amount',                 decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 42, definition: 'Senior Citizen Discount / PWD Discount',   decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 43, definition: 'Total Refund Amount',                      decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 44, definition: 'Total Returned Items Amount',              decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 45, definition: 'Total Other Taxes',                        decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 46, definition: 'Total Service Charge Amount',              decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 47, definition: 'Total Adjustment Discount',                decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 48, definition: 'Total Void Amount',                        decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 49, definition: 'Total Discount Cards',                     decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 50, definition: 'Total Delivery Charges',                   decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 51, definition: 'Total Gift Certificates/Checks Redeemed', decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 52, definition: 'Store Specific Discount 1 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 53, definition: 'Store Specific Discount 2 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 54, definition: 'Store Specific Discount 3 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 55, definition: 'Store Specific Discount 4 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 56, definition: 'Store Specific Discount 5 (Approved)',     decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  /**
   * Line 57 — KNOWN ANOMALY:
   * In the original VBA, totalNASD_NV (sum of lines 58–62) is written to
   * Column D of this row WITHOUT the standard ÷100 division applied to every
   * other currency field. This is replicated faithfully. The field is not
   * validated (no Pass/Failed check). The UI surfaces a ⚠️ warning tooltip.
   */
  { lineItem: 57, definition: 'Total of all Non-Approved Store Discounts',decodeMode: 'currency', isValidated: false, section: 'nonvat', knownAnomaly: 'Known anomaly from original VBA workbook: this admin value was computed without the standard ÷100 division applied to all other currency fields. This field is not validated. Displayed for reference only.' },
  { lineItem: 58, definition: 'Store Specific Discount 1 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 59, definition: 'Store Specific Discount 2 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 60, definition: 'Store Specific Discount 3 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 61, definition: 'Store Specific Discount 4 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 62, definition: 'Store Specific Discount 5 (Not Approved)', decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 63, definition: 'Total VAT/Tax Amount',                     decodeMode: 'currency', isValidated: false, section: 'nonvat' },
  { lineItem: 64, definition: 'Total Net Sales Amount',                   decodeMode: 'currency', isValidated: true,  section: 'nonvat' },
  { lineItem: 65, definition: 'Grand Total Net Sales',                    decodeMode: 'currency', isValidated: true,  section: 'nonvat' },
];

/* ============================================================
   2. FILE DECODER
   Attempts UTF-8 first (strict), falls back to windows-1252
   (covers ANSI/Latin-1 legacy POS files), then iso-8859-1.
   ============================================================ */

/**
 * Decode an ArrayBuffer to a string.
 * @param {ArrayBuffer} buffer
 * @returns {string}
 */
function decodeFileBuffer(buffer) {
  // UTF-8 with fatal=true so we detect encoding errors
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(buffer);
  } catch (_) { /* fall through */ }
  // Windows-1252 covers all ANSI POS files used in the Philippines
  try {
    return new TextDecoder('windows-1252', { fatal: false }).decode(buffer);
  } catch (_) { /* fall through */ }
  // Last resort
  return new TextDecoder('iso-8859-1', { fatal: false }).decode(buffer);
}

/* ============================================================
   3. FILE PARSER
   Parses the .001 fixed-width format into structured line objects
   and populates the accumulator integers needed by the calculator.

   File format per line:
     chars 0-1  : line number (e.g. "01", "34")
     chars 2-13 : 12-char value field (right-padded)
     then CRLF or LF

   Currency decoding: raw integer string → parseFloat → ÷ 100
   Integer decoding:  raw integer string → parseInt  (no division)
   Text decoding:     raw string as-is (trimmed)
   ============================================================ */

/** Round to 2 decimal places — matches VBA Format("#,###,##0.00") */
function round2(n) {
  return Math.round(n * 100) / 100;
}

/**
 * Parse raw .001 file content into lines and accumulators.
 * @param {string} content  Decoded file text
 * @returns {{ lines: Array, acc: Object }}
 */
function parseFile(content) {
  // Normalise line endings: CRLF → LF, lone CR → LF, then split
  const rawLines = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');

  // Strip trailing empty lines (VBA Split on vbCrLf may produce one)
  while (rawLines.length > 0 && rawLines[rawLines.length - 1].trim() === '') {
    rawLines.pop();
  }

  if (rawLines.length < 65) {
    throw new Error(
      `File appears incomplete — expected at least 65 lines, found ${rawLines.length}. ` +
      `Please verify this is a valid POS .001 remittance file.`
    );
  }

  // ── Accumulators (raw integer cents, before ÷100) ──────────────────────
  let oldACCSales      = 0;  // i=3  → line 4
  let totGrossAmount   = 0;  // i=5  → line 6
  let totalDeductions  = 0;  // i=6..21  → lines 7-22
  let totalNASD        = 0;  // i=23..27 → lines 24-28
  let newAccNV         = 0;  // i=36 → line 37 (Old Acc Sales Non-VAT)
  let totGrossNV       = 0;  // i=38 → line 39
  let totalDeductionsNV = 0; // i=40..55 → lines 41-56
  let totalNASD_NV     = 0;  // i=57..61 → lines 58-62

  const lines = [];

  for (let i = 0; i < 65; i++) {
    const lineItem = i + 1;
    const raw = (rawLines[i] || '').substring(2, 14).trim(); // chars 2-13
    const meta = LINE_ITEMS[i]; // LINE_ITEMS is 0-indexed, i corresponds to lineItem-1

    let tenantValue;
    if (meta.decodeMode === 'text') {
      tenantValue = raw;
    } else if (meta.decodeMode === 'integer') {
      tenantValue = parseInt(raw, 10) || 0;
    } else {
      // currency: raw integer ÷ 100
      tenantValue = (parseFloat(raw) || 0) / 100;
    }

    lines.push({ lineItem, rawValue: raw, tenantValue, meta });

    // ── Populate accumulators ───────────────────────────────────────────
    const rawNum = parseFloat(raw) || 0;

    if (i === 3)  oldACCSales     = rawNum;
    if (i === 5)  totGrossAmount  = rawNum;
    if (i >= 7  && i <= 21) totalDeductions  += rawNum;
    if (i >= 23 && i <= 27) totalNASD        += rawNum;
    if (i === 36) newAccNV        = rawNum;
    if (i === 38) totGrossNV      = rawNum;
    if (i >= 40 && i <= 55) totalDeductionsNV += rawNum;
    if (i >= 57 && i <= 61) totalNASD_NV     += rawNum;
  }

  return {
    lines,
    acc: { oldACCSales, totGrossAmount, totalDeductions, totalNASD,
           newAccNV, totGrossNV, totalDeductionsNV, totalNASD_NV },
  };
}

/* ============================================================
   4. CALCULATOR
   Computes all 11 admin-expected Column D values from the raw
   accumulator integers. Faithfully ports the VBA sub
   ReadTextFileAndWriteToColumns from POSFileChecker_v1.xlsm.

   Philippine VAT back-calculation:
     VAT = (GrossAmount − Deductions) × 12 / 112
     Net = (GrossAmount − Deductions) / 1.12

   All inputs are raw integer cents (×100).
   All outputs are rounded to 2dp before returning.
   ============================================================ */

/**
 * @param {Object} acc  Parsed accumulators (raw integer cents)
 * @returns {Object}    Map of lineItem → admin expected value
 */
function computeAdminValues(acc) {
  const { oldACCSales, totGrossAmount, totalDeductions, totalNASD,
          newAccNV, totGrossNV, totalDeductionsNV } = acc;

  // Core VAT intermediates (still in raw cents)
  const vatOnNetSales   = (totGrossAmount - totalDeductions) * 0.12 / 1.12;
  const netSalesExclVAT = totGrossAmount - totalDeductions - vatOnNetSales;
  //   ≡ (totGrossAmount - totalDeductions) / 1.12

  // Core Non-VAT intermediate (raw cents)
  const netSalesNonVAT = totGrossNV - totalDeductionsNV;

  return {
    /**
     * Line 5 — New Accumulated Sales (VAT)
     * = OldAcc/100 + NetSalesExclVAT/100
     */
    5:  round2(oldACCSales / 100 + netSalesExclVAT / 100),

    /**
     * Line 6 — Total Gross Amount (VAT)
     * = totGrossAmount / 100
     * Note: VBA formula strips and re-adds VAT which cancels out —
     * this is a structural identity check; it always equals the
     * tenant's reported gross if the file parses correctly.
     */
    6:  round2(totGrossAmount / 100),

    /**
     * Line 7 — Total Deductions (VAT)
     * = sum of file lines 7–22 / 100
     */
    7:  round2(totalDeductions / 100),

    /**
     * Line 24 — Total Non-Approved Store Discounts (VAT)
     * = sum of file lines 25–29 / 100
     */
    24: round2(totalNASD / 100),

    /**
     * Line 30 — Total VAT/Tax Amount
     * = (Gross − Deductions) × 12/112 / 100
     */
    30: round2(vatOnNetSales / 100),

    /**
     * Line 31 — Total Net Sales Amount (VAT)
     * = (Gross − Deductions) / 1.12 / 100
     */
    31: round2(netSalesExclVAT / 100),

    /**
     * Line 36 — Amount (Non-VAT section header)
     * Same formula as Line 31 — the Non-VAT "Amount" entry is
     * expected to equal the VAT section's net sales.
     */
    36: round2(netSalesExclVAT / 100),

    /**
     * Line 38 — New Accumulated Sales (Non-VAT)
     * = OldAccNV/100 + (GrossNV − DeductionsNV)/100
     */
    38: round2(newAccNV / 100 + netSalesNonVAT / 100),

    /**
     * Line 40 — Total Deductions (Non-VAT)
     * = sum of file lines 41–56 / 100
     */
    40: round2(totalDeductionsNV / 100),

    /**
     * Line 64 — Total Net Sales Amount (Non-VAT)
     * = (GrossNV − DeductionsNV) / 100
     */
    64: round2(netSalesNonVAT / 100),

    /**
     * Line 65 — Grand Total Net Sales
     * = VAT net sales + Non-VAT net sales
     */
    65: round2((netSalesExclVAT + netSalesNonVAT) / 100),
  };
}

/* ============================================================
   5. VALIDATOR
   Orchestrates the full pipeline and builds the result object.

   TOLERANCE RULE:
   A validated field is marked "Failed" only if the absolute
   difference between the tenant-reported value and the
   admin-expected value is STRICTLY GREATER than ±0.01.
   A difference of exactly 0.01 (or less) is treated as a Pass.
   This is the ONLY comparison point in the app that decides
   Pass/Failed status, so the tolerance is centralised here in
   `valuesMatchWithinTolerance()`.
   ============================================================ */

/** Monetary comparison tolerance, in PHP. */
const TOLERANCE = 0.01;

/**
 * Compare two monetary values allowing a ±0.01 tolerance.
 * A tiny epsilon is added to guard against binary floating-point
 * representation artifacts (e.g. 0.010000000000000009).
 * @param {number} a
 * @param {number} b
 * @returns {boolean} true if |a - b| <= 0.01 (within tolerance)
 */
function valuesMatchWithinTolerance(a, b) {
  const diff = Math.abs(round2(a) - round2(b));
  return diff <= TOLERANCE + 1e-9;
}

/**
 * Process an ArrayBuffer from a .001 file upload.
 * @param {ArrayBuffer} buffer
 * @param {string} filename
 * @returns {Object} ValidationResult
 */
function processFile(buffer, filename) {
  const content = decodeFileBuffer(buffer);
  const { lines, acc } = parseFile(content);
  const adminValues = computeAdminValues(acc);

  const rows = LINE_ITEMS.map(meta => {
    const parsed = lines.find(l => l.lineItem === meta.lineItem);
    const tenantValue = parsed ? parsed.tenantValue : null;
    const adminValue  = meta.isValidated ? (adminValues[meta.lineItem] ?? null) : null;

    let status = null;
    if (meta.isValidated && tenantValue !== null && adminValue !== null) {
      // Round both to 2dp, then compare with a ±0.01 tolerance.
      // A difference of exactly 0.01 (or less) counts as Pass;
      // only differences greater than 0.01 are marked Failed.
      const tRounded = round2(typeof tenantValue === 'number' ? tenantValue : 0);
      const aRounded = round2(adminValue);
      status = valuesMatchWithinTolerance(tRounded, aRounded) ? 'Pass' : 'Failed';
    }

    return {
      lineItem:   meta.lineItem,
      definition: meta.definition,
      tenantValue,
      adminValue,
      status,
      isValidated: meta.isValidated,
      section:    meta.section,
      knownAnomaly: meta.knownAnomaly || null,
    };
  });

  const validated  = rows.filter(r => r.isValidated);
  const passed     = validated.filter(r => r.status === 'Pass').length;
  const failed     = validated.filter(r => r.status === 'Failed').length;

  return {
    filename,
    processedAt: new Date(),
    tenantCode:     String(lines[0]?.tenantValue || '').trim() || 'Unknown',
    terminalNumber: String(lines[1]?.tenantValue || '').trim() || 'Unknown',
    posDate:        String(lines[2]?.tenantValue || '').trim() || 'Unknown',
    rows,
    summary: { totalChecked: validated.length, passed, failed, allPassed: failed === 0 },
  };
}

/* ============================================================
   6. FORMATTERS
   ============================================================ */

/**
 * Format a numeric value as PHP currency (2dp, locale-aware).
 * @param {number|string|null} value
 * @returns {string}
 */
function fmtCurrency(value) {
  if (value === null || value === undefined) return '—';
  if (typeof value === 'string') return value;
  return value.toLocaleString('en-PH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/**
 * Format a POS date string (mmddyyyy) for display.
 * @param {string} raw
 * @returns {string}
 */
function fmtDate(raw) {
  if (!raw || raw === 'Unknown') return raw;
  if (/^\d{8}$/.test(raw)) {
    return `${raw.slice(0,2)}/${raw.slice(2,4)}/${raw.slice(4)}`;
  }
  return raw;
}

/**
 * Format a cell value for display in the table.
 * @param {number|string|null} value
 * @param {string} decodeMode
 * @returns {string}
 */
function fmtCellValue(value, decodeMode) {
  if (value === null || value === undefined) return '—';
  if (decodeMode === 'text' || decodeMode === 'integer') return String(value);
  return fmtCurrency(value);
}

/* ============================================================
   7. PDF EXPORT
   Uses jsPDF + jsPDF-AutoTable loaded from CDN in index.html.
   Generates a landscape A4 report matching the spreadsheet
   column layout (A–E). Operates on a single ValidationResult —
   call with whichever file is currently open in the detail view.
   ============================================================ */

/**
 * Export the validation result as a PDF and trigger download.
 * @param {Object} result  ValidationResult
 */
async function exportToPDF(result) {
  // jsPDF and autoTable are loaded globally via <script> in index.html
  if (typeof window.jspdf === 'undefined') {
    alert('PDF export library is loading. Please try again in a moment.');
    return;
  }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

  // ── Header ─────────────────────────────────────────────────────────────
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(14);
  doc.setTextColor(15, 28, 46);
  doc.text('POS File Validation Report', 14, 16);

  doc.setFont('helvetica', 'normal');
  doc.setFontSize(9);
  doc.setTextColor(80, 100, 120);
  doc.text(`File: ${result.filename}`, 14, 23);
  doc.text(`Tenant Code: ${result.tenantCode}`, 14, 28);
  doc.text(`Terminal No.: ${result.terminalNumber}`, 75, 28);
  doc.text(`POS Date: ${fmtDate(result.posDate)}`, 145, 28);
  doc.text(`Processed: ${result.processedAt.toLocaleString('en-PH')}`, 14, 33);

  // ── Summary pills ──────────────────────────────────────────────────────
  const { passed, failed, totalChecked } = result.summary;
  const sy = 38;
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(8);

  doc.setFillColor(0, 200, 150);
  doc.roundedRect(14, sy, 38, 8, 2, 2, 'F');
  doc.setTextColor(255, 255, 255);
  doc.text(`\u2713 ${passed} Passed`, 33, sy + 5.5, { align: 'center' });

  doc.setFillColor(229, 62, 62);
  doc.roundedRect(56, sy, 38, 8, 2, 2, 'F');
  doc.text(`\u2717 ${failed} Failed`, 75, sy + 5.5, { align: 'center' });

  doc.setFillColor(42, 78, 122);
  doc.roundedRect(98, sy, 46, 8, 2, 2, 'F');
  doc.text(`${totalChecked} Checks Total`, 121, sy + 5.5, { align: 'center' });

  // ── Table ──────────────────────────────────────────────────────────────
  const tableData = result.rows.map(row => [
    String(row.lineItem),
    row.definition,
    row.tenantValue !== null ? fmtCellValue(row.tenantValue, LINE_ITEMS[row.lineItem - 1].decodeMode) : '—',
    row.adminValue  !== null ? fmtCurrency(row.adminValue)  : '—',
    row.status || '—',
  ]);

  doc.autoTable({
    startY: sy + 12,
    head: [['#', 'Line Item Definition', 'Tenant File Value (C)', 'Admin Expected (D)', 'Status (E)']],
    body: tableData,
    styles: { font: 'helvetica', fontSize: 7.5, cellPadding: 2.5, textColor: [30, 50, 70] },
    headStyles: { fillColor: [15, 28, 46], textColor: [255, 255, 255], fontStyle: 'bold', fontSize: 8 },
    columnStyles: {
      0: { cellWidth: 10, halign: 'center' },
      1: { cellWidth: 92 },
      2: { cellWidth: 42, halign: 'right', font: 'courier' },
      3: { cellWidth: 42, halign: 'right', font: 'courier' },
      4: { cellWidth: 22, halign: 'center' },
    },
    didDrawCell(data) {
      if (data.column.index === 4 && data.section === 'body') {
        if (data.cell.raw === 'Pass')   doc.setTextColor(0, 160, 100);
        else if (data.cell.raw === 'Failed') doc.setTextColor(220, 50, 50);
        else doc.setTextColor(120, 140, 160);
      }
    },
    alternateRowStyles: { fillColor: [245, 248, 252] },
    margin: { left: 14, right: 14 },
  });

  // ── Footer ─────────────────────────────────────────────────────────────
  const pages = doc.internal.getNumberOfPages();
  for (let p = 1; p <= pages; p++) {
    doc.setPage(p);
    doc.setFont('helvetica', 'normal');
    doc.setFontSize(7);
    doc.setTextColor(140, 150, 160);
    doc.text(
      `POS File Checker  |  Page ${p} of ${pages}`,
      doc.internal.pageSize.width / 2,
      doc.internal.pageSize.height - 6,
      { align: 'center' }
    );
  }

  doc.save(`validation_${result.filename.replace(/[^a-zA-Z0-9._-]/g, '_')}.pdf`);
}

/* ============================================================
   8. EXCEL EXPORT
   Uses SheetJS (XLSX) loaded from CDN in index.html.
   Produces a two-sheet workbook: Summary + Validation data,
   mirroring the original spreadsheet's column structure.
   Operates on a single ValidationResult — call with whichever
   file is currently open in the detail view.
   ============================================================ */

/**
 * Export the validation result as an .xlsx file and trigger download.
 * @param {Object} result  ValidationResult
 */
function exportToExcel(result) {
  if (typeof window.XLSX === 'undefined') {
    alert('Excel export library is loading. Please try again in a moment.');
    return;
  }
  const XLSX = window.XLSX;
  const wb = XLSX.utils.book_new();

  // ── Summary sheet ───────────────────────────────────────────────────────
  const summaryData = [
    ['POS File Validation Report'],
    [],
    ['File',            result.filename],
    ['Tenant Code',     result.tenantCode],
    ['Terminal Number', result.terminalNumber],
    ['POS Date',        fmtDate(result.posDate)],
    ['Processed At',    result.processedAt.toLocaleString('en-PH')],
    [],
    ['Checks Passed',   result.summary.passed],
    ['Checks Failed',   result.summary.failed],
    ['Total Checks',    result.summary.totalChecked],
    ['Overall Result',  result.summary.allPassed ? 'ALL PASS' : `${result.summary.failed} FAILED`],
  ];
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
  summarySheet['!cols'] = [{ wch: 20 }, { wch: 42 }];
  XLSX.utils.book_append_sheet(wb, summarySheet, 'Summary');

  // ── Validation sheet (mirrors Columns A–E of the spreadsheet) ────────────
  const headers = [
    'Line Item #',
    'Line Item Definition',
    'Tenant File Value',   // Column C
    'Admin Expected Value',// Column D
    'Status',              // Column E
  ];
  const dataRows = result.rows.map(row => [
    row.lineItem,
    row.definition,
    row.tenantValue !== null
      ? (typeof row.tenantValue === 'number' ? row.tenantValue : String(row.tenantValue))
      : null,
    row.adminValue !== null ? row.adminValue : null,
    row.status || '',
  ]);
  const valSheet = XLSX.utils.aoa_to_sheet([headers, ...dataRows]);
  valSheet['!cols'] = [{ wch: 12 }, { wch: 46 }, { wch: 20 }, { wch: 22 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, valSheet, 'Validation');

  XLSX.writeFile(wb, `validation_${result.filename.replace(/[^a-zA-Z0-9._-]/g, '_')}.xlsx`);
}

/**
 * Export a batch summary (all processed files) as an .xlsx file.
 * Lists File Name / Status / Failed Checks / Total Checked for
 * every file currently in `allResults`.
 * @param {Array} results  Array of ValidationResult
 */
function exportBatchSummaryToExcel(results) {
  if (typeof window.XLSX === 'undefined') {
    alert('Excel export library is loading. Please try again in a moment.');
    return;
  }
  const XLSX = window.XLSX;
  const wb = XLSX.utils.book_new();

  const totalFiles = results.length;
  const totalPass  = results.filter(r => r.summary.allPassed).length;
  const totalFail  = totalFiles - totalPass;

  const summaryData = [
    ['POS File Checker — Batch Summary'],
    [],
    ['Total Files Processed', totalFiles],
    ['Total PASS',            totalPass],
    ['Total FAIL',            totalFail],
    [],
  ];
  const summarySheet = XLSX.utils.aoa_to_sheet(summaryData);
  summarySheet['!cols'] = [{ wch: 26 }, { wch: 20 }];
  XLSX.utils.book_append_sheet(wb, summarySheet, 'Batch Summary');

  const headers = ['File Name', 'Status', 'Failed Checks', 'Total Checked', 'Tenant Code', 'Terminal Number'];
  const dataRows = results.map(r => [
    r.filename,
    r.summary.allPassed ? 'PASS' : 'FAIL',
    r.summary.failed,
    r.summary.totalChecked,
    r.tenantCode,
    r.terminalNumber,
  ]);
  const listSheet = XLSX.utils.aoa_to_sheet([headers, ...dataRows]);
  listSheet['!cols'] = [{ wch: 30 }, { wch: 10 }, { wch: 14 }, { wch: 14 }, { wch: 16 }, { wch: 18 }];
  XLSX.utils.book_append_sheet(wb, listSheet, 'Files');

  XLSX.writeFile(wb, `validation_batch_summary_${Date.now()}.xlsx`);
}

/* ============================================================
   9. UI RENDERING
   ============================================================ */

/** Currently active filter (within the single-file detail view): 'all' | 'validated' | 'failed' */
let activeFilter = 'all';

/** All processed files in the current batch — each is a ValidationResult with an added `id`. */
let allResults = [];

/** The ValidationResult currently open in the detail view (or null). */
let currentResult = null;

/** Monotonically increasing id counter for tagging results. */
let idCounter = 0;

/** Generate a unique id for a newly processed file's result. */
function generateId() {
  idCounter += 1;
  return `file_${idCounter}_${Date.now()}`;
}

// ── DOM references ────────────────────────────────────────────────────────
const uploadView   = document.getElementById('upload-view');
const summaryView  = document.getElementById('summary-view');
const resultsView  = document.getElementById('results-view');
const dropZone     = document.getElementById('drop-zone');
const fileInput    = document.getElementById('file-input');
const dropTextPrimary   = document.getElementById('drop-text-primary');
const dropTextSecondary = document.getElementById('drop-text-secondary');
const dropIconBox  = document.getElementById('drop-icon-box');
const errorBanner  = document.getElementById('error-banner');
const errorMsg     = document.getElementById('error-msg');

// Batch / summary view elements
const batchBanner       = document.getElementById('batch-result-banner');
const batchStatPass     = document.getElementById('batch-stat-pass');
const batchStatFail     = document.getElementById('batch-stat-fail');
const batchStatTotal    = document.getElementById('batch-stat-total');
const summaryTableBody  = document.getElementById('summary-table-body');
const summarySubtitle   = document.getElementById('summary-subtitle');
const btnAddMore        = document.getElementById('btn-add-more');
const btnClearAllBatch  = document.getElementById('btn-clear-all');
const btnExportBatch    = document.getElementById('btn-export-batch');

// Detail (single-file) results elements
const resultBanner    = document.getElementById('result-banner');
const metaPills       = document.getElementById('meta-pills');
const statPassed      = document.getElementById('stat-passed');
const statFailed      = document.getElementById('stat-failed');
const statTotal       = document.getElementById('stat-total');
const filterTabs      = document.querySelectorAll('.filter-tab');
const tabCountAll     = document.getElementById('tab-count-all');
const tabCountVal     = document.getElementById('tab-count-val');
const tabCountFail    = document.getElementById('tab-count-fail');
const tableBody       = document.getElementById('table-body');
const btnExportPDF    = document.getElementById('btn-export-pdf');
const btnExportExcel  = document.getElementById('btn-export-excel');
const resultsSubtitle = document.getElementById('results-subtitle');
const btnBackToSummary = document.getElementById('btn-back-to-summary');
const btnReset          = document.getElementById('btn-reset');

/** Show an error message and ensure the upload view is visible. */
function showError(msg) {
  errorMsg.textContent = msg;
  errorBanner.classList.add('visible');
}

/** Hide the error banner. */
function hideError() {
  errorBanner.classList.remove('visible');
}

/** Switch to processing state in the drop zone. @param {boolean} on @param {number} [count] */
function setProcessing(on, count) {
  if (on) {
    dropZone.classList.add('processing');
    dropIconBox.innerHTML = '<div class="drop-spinner"></div>';
    dropTextPrimary.textContent   = count && count > 1
      ? `Processing ${count} files\u2026`
      : 'Processing file\u2026';
    dropTextSecondary.textContent = 'Parsing and running validation checks';
  } else {
    dropZone.classList.remove('processing');
    dropIconBox.innerHTML = iconUpload();
    dropTextPrimary.innerHTML     = 'Drop your <span>.001&nbsp;&ndash;&nbsp;.999</span> file(s) here';
    dropTextSecondary.textContent = 'or browse to select — multiple files supported';
  }
}

/** Render summary stats and the result banner for the single open file (detail view). */
function renderSummary(result) {
  const { passed, failed, totalChecked, allPassed } = result.summary;

  // Banner
  if (allPassed) {
    resultBanner.className = 'result-banner all-pass';
    resultBanner.innerHTML = `${iconCheckCircle()}<span>All ${totalChecked} checks passed — file is valid.</span>`;
  } else {
    resultBanner.className = 'result-banner has-fail';
    resultBanner.innerHTML = `${iconXCircle()}<span>${failed} of ${totalChecked} checks failed — review highlighted rows below.</span>`;
  }

  // Meta pills
  metaPills.innerHTML = [
    metaPill('File',     result.filename,              true),
    metaPill('Tenant',   result.tenantCode,             false),
    metaPill('Terminal', result.terminalNumber,         false),
    metaPill('Date',     fmtDate(result.posDate),       false),
  ].join('');

  // Stat cards
  statPassed.textContent = passed;
  statFailed.textContent = failed;
  statTotal.textContent  = totalChecked;

  // Filter tab counts
  tabCountAll.textContent  = result.rows.length;
  tabCountVal.textContent  = result.rows.filter(r => r.isValidated).length;
  tabCountFail.textContent = failed;
  tabCountFail.className   = 'tab-count' + (failed > 0 ? ' fail-count' : '');

  // Subtitle
  resultsSubtitle.textContent =
    `65 line items · 11 validated fields · ±0.01 tolerance · processed ${result.processedAt.toLocaleTimeString('en-PH')}`;
}

/** Create a meta pill HTML string. */
function metaPill(label, value, mono) {
  return `<span class="meta-pill">
    <span class="label">${esc(label)}:</span>
    <span class="value${mono ? ' mono' : ''}">${esc(value)}</span>
  </span>`;
}

/** Render the validation table rows (within the detail view) for the current filter. */
function renderTable() {
  if (!currentResult) return;

  const rows = currentResult.rows.filter(row => {
    if (activeFilter === 'validated') return row.isValidated;
    if (activeFilter === 'failed')    return row.status === 'Failed';
    return true;
  });

  if (rows.length === 0) {
    tableBody.innerHTML = `<tr><td colspan="5" class="table-empty">No rows match the selected filter.</td></tr>`;
    return;
  }

  let html = '';
  let prevSection = null;

  rows.forEach(row => {
    // Section separator before Non-VAT section (only in 'all' view)
    if (activeFilter === 'all' && row.section === 'nonvat' && prevSection === 'vat') {
      html += `<tr class="section-sep">
        <td colspan="5">
          <div class="section-sep-inner">
            <div class="section-sep-line"></div>
            <span class="section-sep-label">Non-VAT Sales Section</span>
            <div class="section-sep-line"></div>
          </div>
        </td>
      </tr>`;
    }
    prevSection = row.section;

    // Determine stripe class
    const stripe = row.status === 'Pass' ? 'stripe-pass'
                 : row.status === 'Failed' ? 'stripe-fail'
                 : 'stripe-none';

    // Tenant value display
    const tMeta = LINE_ITEMS[row.lineItem - 1];
    const tDisplay = row.tenantValue !== null ? fmtCellValue(row.tenantValue, tMeta.decodeMode) : '—';
    const aDisplay = row.adminValue  !== null ? fmtCurrency(row.adminValue)  : '—';
    const tMuted   = tDisplay === '—';
    const aMuted   = aDisplay === '—';

    // Status cell
    let statusCell;
    if (row.status === 'Pass') {
      statusCell = `<span class="badge pass">Pass</span>`;
    } else if (row.status === 'Failed') {
      statusCell = `<button class="btn-expand" id="expand-${row.lineItem}" aria-expanded="false" aria-controls="detail-${row.lineItem}" onclick="toggleDetail(${row.lineItem})">
        <span class="badge fail">Failed</span>
        ${iconChevronDown()}
      </button>`;
    } else {
      statusCell = `<span style="color:var(--slate-muted);font-size:.75rem">—</span>`;
    }

    // Anomaly badge
    let anomalyBadge = '';
    if (row.knownAnomaly) {
      anomalyBadge = `<span class="anomaly-wrap">
        <button class="anomaly-btn" tabindex="0" aria-label="Known anomaly">${iconWarning()}</button>
        <div class="anomaly-tip"><p class="anomaly-tip-title">Known Anomaly</p>${esc(row.knownAnomaly)}</div>
      </span>`;
    }

    html += `<tr class="${stripe}" id="row-${row.lineItem}">
      <td class="td-num">${row.lineItem}</td>
      <td class="td-def${row.isValidated ? ' validated' : ''}">
        <div class="td-def-inner">${esc(row.definition)}${anomalyBadge}</div>
      </td>
      <td class="td-val${row.isValidated ? ' validated' : tMuted ? ' muted' : ''}">${tDisplay}</td>
      <td class="td-val${aMuted ? ' muted' : ''}">${aDisplay}</td>
      <td class="td-status">${statusCell}</td>
    </tr>`;

    // Detail tray for Failed rows
    if (row.status === 'Failed') {
      const tenantNum = typeof row.tenantValue === 'number' ? row.tenantValue : 0;
      const adminNum  = typeof row.adminValue  === 'number' ? row.adminValue  : 0;
      const diff      = tenantNum - adminNum;
      html += `<tr class="detail-row" id="detail-${row.lineItem}">
        <td colspan="5">
          <div class="detail-inner">
            <p class="detail-title">Discrepancy Detail</p>
            <div class="detail-cards">
              <div class="detail-card">
                <div class="detail-card-label">Tenant reported</div>
                <div class="detail-card-value">${fmtCellValue(row.tenantValue, tMeta.decodeMode)}</div>
              </div>
              <div class="detail-card">
                <div class="detail-card-label">Admin expected</div>
                <div class="detail-card-value">${fmtCurrency(row.adminValue)}</div>
              </div>
            </div>
            <p class="detail-diff">Difference: <span>${fmtCurrency(diff)}</span> (tolerance: ±${TOLERANCE.toFixed(2)})</p>
          </div>
        </td>
      </tr>`;
    }
  });

  tableBody.innerHTML = html;
}

/** Toggle the detail tray for a Failed row. */
function toggleDetail(lineItem) {
  const btn    = document.getElementById(`expand-${lineItem}`);
  const detail = document.getElementById(`detail-${lineItem}`);
  if (!btn || !detail) return;
  const open = detail.classList.toggle('open');
  btn.classList.toggle('open', open);
  btn.setAttribute('aria-expanded', String(open));
}
// Expose globally for inline onclick handlers
window.toggleDetail = toggleDetail;

/** Update which filter tab is active and re-render the detail table. */
function updateFilterTabs() {
  filterTabs.forEach(tab => {
    tab.classList.toggle('active', tab.dataset.filter === activeFilter);
  });
  renderTable();
}

/* ── View switching ────────────────────────────────────────────────────── */

/** Show the upload view and clear ALL batch state. */
function showUpload() {
  allResults = [];
  currentResult = null;
  activeFilter = 'all';
  summaryView.classList.remove('visible');
  resultsView.classList.remove('visible');
  uploadView.style.display = '';
  setProcessing(false);
  hideError();
  fileInput.value = '';
}

/** Render and show the batch summary view (list of all processed files). */
function showSummaryView() {
  uploadView.style.display = 'none';
  resultsView.classList.remove('visible');
  renderBatchSummary();
  summaryView.classList.add('visible');
  hideError();
}

/** Open the detail (single-file) view for a specific result id. */
function openDetail(id) {
  const result = allResults.find(r => String(r.id) === String(id));
  if (!result) return;
  currentResult = result;
  activeFilter = 'all';
  uploadView.style.display = 'none';
  summaryView.classList.remove('visible');
  renderSummary(result);
  updateFilterTabs(); // also renders the table
  resultsView.classList.add('visible');
  hideError();
}

/** Render the aggregate stat cards + per-file table in the summary view. */
function renderBatchSummary() {
  const totalFiles = allResults.length;
  const totalPass  = allResults.filter(r => r.summary.allPassed).length;
  const totalFail  = totalFiles - totalPass;

  if (batchBanner) {
    if (totalFiles === 0) {
      batchBanner.className = 'result-banner';
      batchBanner.innerHTML = '';
    } else if (totalFail === 0) {
      batchBanner.className = 'result-banner all-pass';
      batchBanner.innerHTML = `${iconCheckCircle()}<span>All ${totalFiles} file${totalFiles === 1 ? '' : 's'} passed validation.</span>`;
    } else {
      batchBanner.className = 'result-banner has-fail';
      batchBanner.innerHTML = `${iconXCircle()}<span>${totalFail} of ${totalFiles} file${totalFiles === 1 ? '' : 's'} failed validation — click a row below to view details.</span>`;
    }
  }

  if (batchStatTotal) batchStatTotal.textContent = totalFiles;
  if (batchStatPass)  batchStatPass.textContent  = totalPass;
  if (batchStatFail)  batchStatFail.textContent  = totalFail;

  if (summarySubtitle) {
    summarySubtitle.textContent = `${totalFiles} file${totalFiles === 1 ? '' : 's'} processed · click a row to view its detailed report`;
  }

  renderSummaryTable();
}

/** Render the clickable per-file summary table. */
function renderSummaryTable() {
  if (!summaryTableBody) return;

  if (allResults.length === 0) {
    summaryTableBody.innerHTML = `<tr><td colspan="5" class="table-empty">No files processed yet.</td></tr>`;
    return;
  }

  let html = '';
  allResults.forEach((r, idx) => {
    const isPass = r.summary.allPassed;
    const stripe = isPass ? 'stripe-pass' : 'stripe-fail';
    html += `<tr class="${stripe} summary-row" id="summary-row-${r.id}" data-id="${esc(r.id)}" tabindex="0" role="button" aria-label="View detailed report for ${esc(r.filename)}">
      <td class="td-num">${idx + 1}</td>
      <td class="td-def validated"><div class="td-def-inner">${esc(r.filename)}</div></td>
      <td class="td-status"><span class="badge ${isPass ? 'pass' : 'fail'}">${isPass ? 'PASS' : 'FAIL'}</span></td>
      <td class="td-val validated">${r.summary.failed}</td>
      <td class="td-status">${iconChevronRight()}</td>
    </tr>`;
  });
  summaryTableBody.innerHTML = html;

  // Wire row clicks / keyboard activation
  summaryTableBody.querySelectorAll('tr.summary-row').forEach(row => {
    row.addEventListener('click', () => openDetail(row.dataset.id));
    row.addEventListener('keydown', e => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        openDetail(row.dataset.id);
      }
    });
  });
}

/* ============================================================
   10. EVENT WIRING
   ============================================================ */

/**
 * Return true if the filename has a valid POS numeric extension.
 *
 * Rules (per spec):
 *   - Extension must be present (filename must contain a dot)
 *   - Extension must be EXACTLY 3 characters after the dot
 *   - All 3 characters must be ASCII digits (0-9)
 *   - Valid range: .001 – .999  (.000 is technically accepted by the
 *     regex but no POS terminal produces it; excluded in the error msg)
 *
 * Examples:
 *   Accepted : sales.001  sales.015  sales.100  sales.999  DATA.042
 *   Rejected : sales.txt  sales.csv  sales.1    sales.1000
 *              sales.abc  sales.00A  sales.     noextension
 *
 * @param {string} filename
 * @returns {boolean}
 */
function isValidPosExtension(filename) {
  // Extract everything after the last dot (case-insensitive not needed —
  // digits have no case, but we normalise anyway for safety)
  const dot = filename.lastIndexOf('.');
  if (dot === -1) return false;            // no extension at all
  const ext = filename.slice(dot + 1);    // e.g. "001", "txt", "1000"
  return /^\d{3}$/.test(ext);             // exactly 3 decimal digits
}

/**
 * Read a File as an ArrayBuffer, wrapped in a Promise.
 * @param {File} file
 * @returns {Promise<ArrayBuffer>}
 */
function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload  = () => resolve(reader.result);
    reader.onerror = () => reject(new Error('Could not read the file.'));
    reader.readAsArrayBuffer(file);
  });
}

/**
 * Process a single File object end-to-end.
 * @param {File} file
 * @returns {Promise<{ok: true, result: Object} | {ok: false, filename: string, message: string}>}
 */
async function processSingleFile(file) {
  if (!isValidPosExtension(file.name)) {
    const dot = file.name.lastIndexOf('.');
    const ext = dot !== -1 ? file.name.slice(dot) : '(no extension)';
    return {
      ok: false,
      filename: file.name,
      message: `Invalid file type — POS remittance files must have a 3-digit numeric extension (.001 – .999). Received: "${ext}"`,
    };
  }
  try {
    const buffer = await readFileAsArrayBuffer(file);
    const result = processFile(buffer, file.name);
    result.id = generateId();
    return { ok: true, result };
  } catch (err) {
    return { ok: false, filename: file.name, message: err.message || 'An unexpected error occurred while processing the file.' };
  }
}

/**
 * Handle a batch of selected/dropped files: process every one of them,
 * append successful results to `allResults`, and surface any per-file
 * errors without blocking the rest of the batch.
 * @param {FileList|File[]} fileList
 */
async function handleFiles(fileList) {
  const files = Array.from(fileList || []);
  if (files.length === 0) return;

  hideError();

  // Show the spinner via the upload-view drop zone regardless of which
  // view triggered this (e.g. "Add More Files" from the summary view).
  uploadView.style.display = '';
  summaryView.classList.remove('visible');
  resultsView.classList.remove('visible');
  setProcessing(true, files.length);

  const newResults = [];
  const errorMessages = [];

  for (const file of files) {
    const outcome = await processSingleFile(file);
    if (outcome.ok) {
      newResults.push(outcome.result);
    } else {
      errorMessages.push(`"${outcome.filename}": ${outcome.message}`);
    }
  }

  setProcessing(false);

  if (newResults.length > 0) {
    allResults = allResults.concat(newResults);
  }

  if (errorMessages.length > 0) {
    showError(errorMessages.join('  ·  '));
  }

  if (allResults.length > 0) {
    showSummaryView();
    if (errorMessages.length > 0) {
      // Re-assert the error banner since showSummaryView() hides it via hideError().
      showError(errorMessages.join('  ·  '));
    }
  } else {
    // Nothing succeeded — stay on the upload view so the error is visible.
    uploadView.style.display = '';
    fileInput.value = '';
  }
}

// ── Drop zone ─────────────────────────────────────────────────────────────
dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const files = e.dataTransfer?.files;
  if (files && files.length > 0) handleFiles(files);
});
dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('keydown', e => {
  if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fileInput.click(); }
});

// ── File input ────────────────────────────────────────────────────────────
fileInput.addEventListener('change', () => {
  const files = fileInput.files;
  if (files && files.length > 0) handleFiles(files);
});

// ── Error banner close ────────────────────────────────────────────────────
document.getElementById('error-close').addEventListener('click', hideError);

// ── Detail view navigation ────────────────────────────────────────────────
if (btnBackToSummary) {
  btnBackToSummary.addEventListener('click', () => {
    currentResult = null;
    showSummaryView();
  });
}
if (btnReset) {
  // "Clear All Files" — full reset back to the upload view.
  btnReset.addEventListener('click', showUpload);
}

// ── Summary view controls ─────────────────────────────────────────────────
if (btnAddMore) {
  btnAddMore.addEventListener('click', () => fileInput.click());
}
if (btnClearAllBatch) {
  btnClearAllBatch.addEventListener('click', showUpload);
}
if (btnExportBatch) {
  btnExportBatch.addEventListener('click', () => {
    if (allResults.length === 0) return;
    btnExportBatch.disabled = true;
    try {
      exportBatchSummaryToExcel(allResults);
    } finally {
      btnExportBatch.disabled = false;
    }
  });
}

// ── Filter tabs (detail view) ─────────────────────────────────────────────
filterTabs.forEach(tab => {
  tab.addEventListener('click', () => {
    activeFilter = tab.dataset.filter;
    updateFilterTabs();
  });
});

// ── Export buttons (detail view — operate on currently open file) ────────
btnExportPDF.addEventListener('click', async () => {
  if (!currentResult) return;
  btnExportPDF.disabled = true;
  try {
    await exportToPDF(currentResult);
  } finally {
    btnExportPDF.disabled = false;
  }
});

btnExportExcel.addEventListener('click', () => {
  if (!currentResult) return;
  btnExportExcel.disabled = true;
  try {
    exportToExcel(currentResult);
  } finally {
    btnExportExcel.disabled = false;
  }
});

/* ============================================================
   INLINE SVG HELPERS
   Returns SVG strings so no external icon font is needed.
   ============================================================ */
function iconUpload() {
  return `<svg viewBox="0 0 24 24"><polyline points="16 16 12 12 8 16"/><line x1="12" y1="12" x2="12" y2="21"/><path d="M20.39 18.39A5 5 0 0 0 18 9h-1.26A8 8 0 1 0 3 16.3"/></svg>`;
}
function iconCheckCircle() {
  return `<svg viewBox="0 0 24 24" style="width:20px;height:20px;stroke:currentColor;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;flex-shrink:0"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/><polyline points="22 4 12 14.01 9 11.01"/></svg>`;
}
function iconXCircle() {
  return `<svg viewBox="0 0 24 24" style="width:20px;height:20px;stroke:currentColor;fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round;flex-shrink:0"><circle cx="12" cy="12" r="10"/><line x1="15" y1="9" x2="9" y2="15"/><line x1="9" y1="9" x2="15" y2="15"/></svg>`;
}
function iconChevronDown() {
  return `<svg viewBox="0 0 24 24"><polyline points="6 9 12 15 18 9"/></svg>`;
}
function iconChevronRight() {
  return `<svg viewBox="0 0 24 24" style="width:16px;height:16px;stroke:var(--slate-muted);fill:none;stroke-width:2;stroke-linecap:round;stroke-linejoin:round"><polyline points="9 18 15 12 9 6"/></svg>`;
}
function iconWarning() {
  return `<svg viewBox="0 0 24 24"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>`;
}
function iconReset() {
  return `<svg viewBox="0 0 24 24"><polyline points="1 4 1 10 7 10"/><path d="M3.51 15a9 9 0 1 0 .49-4.95"/></svg>`;
}
function iconBack() {
  return `<svg viewBox="0 0 24 24"><polyline points="15 18 9 12 15 6"/></svg>`;
}
function iconPlus() {
  return `<svg viewBox="0 0 24 24"><line x1="12" y1="5" x2="12" y2="19"/><line x1="5" y1="12" x2="19" y2="12"/></svg>`;
}

/** HTML-escape a string for safe insertion. */
function esc(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// Initial icon render
dropIconBox.innerHTML = iconUpload();
