/**
 * EXTRACT: one script for both source workbooks (Office Script / TypeScript)
 * ==========================================================================
 * Run it in Power Automate ("Run script") on EITHER source workbook:
 *   - Functional DMB Review Sheets   -> rows with "functionName"
 *   - Strategic Execution Dashboard  -> rows with "strategicImperative"
 * It recognises the workbook by itself (a visible 'AOP Critical' tab means
 * Strategic Execution, otherwise the numbered function tabs are read), so the
 * same script can be picked for both "Run script" steps.
 * Returns: JSON text for sync_masterfile.ts.
 */

type Cell = string | number | boolean;

function main(workbook: ExcelScript.Workbook): string {
  const aop = workbook.getWorksheet("AOP Critical");
  if (aop) {
    const rows = extractStrategic(aop);
    console.log(`Strategic Execution workbook: extracted ${rows.length} KPI rows.`);
    return JSON.stringify(rows);
  }
  const rows = extractFunctional(workbook);
  console.log(`Functional DMB workbook: extracted ${rows.length} KPI rows.`);
  return JSON.stringify(rows);
}

// =====================================================================
// SHARED
// =====================================================================

interface KPIRow {
  functionName?: string;
  strategicImperative?: string;
  sourceSheet: string;
  kpiName: string;
  owner?: string;
  definition: string;
  operator?: string;
  target2026: Cell;
  column1?: string;
  column2?: string;
  unit: string;
  nature: string;
  frequency: string;
  rowType: string;
  jan: Cell; feb: Cell; mar: Cell; apr: Cell; may: Cell; jun: Cell;
  jul: Cell; aug: Cell; sep: Cell; oct: Cell; nov: Cell; dec: Cell;
  dataType?: string;
  category?: string;
  /** Number format of target2026 and of Jan..Dec as in the source ("" = none). */
  target2026Format: string;
  formats: string[];
}

/** A source cell: its value and the number format Excel shows it with. */
interface SourceCell {
  value: Cell;
  format: string;
}

const MONTH_PREFIXES = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];

function dataTypeForUnit(unit: string): string {
  const lower = unit.trim().toLowerCase();
  if (lower === "%") return "Percentage";
  if (lower === "mn" || lower === "k" || lower.indexOf("decimal") >= 0) return "Decimal";
  return "Whole Number";
}

const NO_CELL: SourceCell = { value: "", format: "" };

/**
 * A source cell exactly as Excel shows it: the value with its own number
 * format, so the Masterfile shows 92% as 92%, 92.00% as 92.00% and 91 as 91.
 * The only change: a percentage typed as text ("90%%", "90%") becomes 0.9
 * shown as 90%, with a single % sign.
 */
function sourceCell(raw: Cell | undefined | null, format: string | undefined): SourceCell {
  if (raw === null || raw === undefined) return NO_CELL;
  const text = String(raw).trim();
  if (text === "" || text === "#REF!" || text === "None") return NO_CELL;
  if (typeof raw !== "string") return { value: raw, format: format || "General" };

  const percent = text.match(/^(-?\d+(?:\.(\d+))?)\s*%+$/);
  if (percent) {
    const decimals = Math.min(2, percent[2] ? percent[2].length : 0);
    return {
      value: Number((Number(percent[1]) / 100).toPrecision(12)),
      format: decimals === 0 ? "0%" : (decimals === 1 ? "0.0%" : "0.00%")
    };
  }
  // Other text keeps the Masterfile cell's format.
  return { value: text.indexOf("%%") >= 0 ? text.replace(/%{2,}/g, "%") : raw, format: "" };
}

function readMonthValues(row: Cell[], formats: string[], monthCols: { [monthIndex: number]: number }): SourceCell[] {
  const monthly: SourceCell[] = [];
  for (let m = 0; m < 12; m++) {
    const c = monthCols[m];
    monthly.push(c !== undefined && c < row.length ? sourceCell(row[c], formats[c]) : NO_CELL);
  }
  return monthly;
}

function withMonths(base: KPIRow, monthly: SourceCell[]): KPIRow {
  const v = monthly.map(cell => cell.value);
  base.jan = v[0]; base.feb = v[1]; base.mar = v[2]; base.apr = v[3];
  base.may = v[4]; base.jun = v[5]; base.jul = v[6]; base.aug = v[7];
  base.sep = v[8]; base.oct = v[9]; base.nov = v[10]; base.dec = v[11];
  base.formats = monthly.map(cell => cell.format);
  return base;
}

// =====================================================================
// FUNCTIONAL DMB REVIEW SHEETS
// =====================================================================

/** Tabs that never hold current KPI data. */
const SKIP_SHEETS = ["mpr review", "mastersheet", "sheet1 (2)", "marketing june"];

/** Everything below these headings is RCA / action-tracker content. */
const STOP_HEADINGS = [
  "trends", "root cause analysis", "paretos", "action tracker",
  "if the kpi is reported", "handshake"
];

const FUNCTION_KEYWORDS = [
  "quality", "regulatory", "isc", "procurement", "r&d", "marketing",
  "customer service", "commercial excellence", "nar", "europe", "growth", "finance"
];

interface KPIBlock {
  kpiName: string;
  definition: string;
  operator: string;
  target2026: SourceCell;
  unit: string;
  nature: string;
  frequency: string;
  category: string;
  isSubKpi: boolean;
  groupIndex: number;
  targetValues: SourceCell[];
  actualValues: SourceCell[];
}

function extractFunctional(workbook: ExcelScript.Workbook): KPIRow[] {
  const rows: KPIRow[] = [];

  for (const sheet of workbook.getWorksheets()) {
    const sheetName = sheet.getName().trim();
    const lower = sheetName.toLowerCase();

    // Only visible function tabs; hidden copies hold stale or #REF! data.
    if (sheet.getVisibility() !== ExcelScript.SheetVisibility.visible) continue;
    if (SKIP_SHEETS.indexOf(lower) >= 0) continue;
    if (!FUNCTION_KEYWORDS.some(k => lower.indexOf(k) >= 0)) continue;

    const used = sheet.getUsedRange(true);
    if (!used) continue;
    // Read from A1 so row/column positions match the sheet.
    const range = sheet.getRangeByIndexes(0, 0,
      used.getRowIndex() + used.getRowCount(),
      Math.max(20, used.getColumnIndex() + used.getColumnCount()));
    const values = range.getValues() as Cell[][];
    const formats = range.getNumberFormats() as string[][];
    if (values.length < 5) continue;

    let headerRow = 4;
    for (let r = 0; r < Math.min(10, values.length); r++) {
      const first = String(values[r][0] || "").trim().toLowerCase();
      if (first === "kpi name" || first.indexOf("mandatory kpi") === 0) {
        headerRow = r;
        break;
      }
    }

    const functionName = cleanFunctionName(sheetName);
    const blocks = readKPIBlocks(values, formats, headerRow, functionalMonthColumns(values[headerRow]));
    nameSubKpis(blocks);

    for (const block of blocks) {
      for (const isTarget of [true, false]) {
        rows.push(withMonths({
          functionName: functionName,
          sourceSheet: sheetName,
          kpiName: block.kpiName,
          definition: block.definition,
          operator: block.operator,
          target2026: block.target2026.value,
          unit: block.unit,
          nature: block.nature,
          frequency: block.frequency,
          rowType: isTarget ? "Target" : "Actual",
          jan: "", feb: "", mar: "", apr: "", may: "", jun: "",
          jul: "", aug: "", sep: "", oct: "", nov: "", dec: "",
          dataType: dataTypeForUnit(block.unit),
          category: block.category,
          target2026Format: block.target2026.format,
          formats: []
        }, isTarget ? block.targetValues : block.actualValues));
      }
    }
  }
  return rows;
}

/**
 * Read the KPI table of one function tab. A KPI row carries its name in
 * column A; a sub-KPI (e.g. the Z30 and Z90 complaint rates) leaves column A
 * empty and names itself in the definition.
 */
function readKPIBlocks(
  values: Cell[][],
  formats: string[][],
  headerRow: number,
  monthCols: { [m: number]: number }
): KPIBlock[] {
  const blocks: KPIBlock[] = [];
  let category = "";
  let groupIndex = -1;
  let current: KPIBlock | null = null;

  // The first "Mandatory Outcome KPI's" heading sits above the header row.
  for (let r = 0; r < headerRow; r++) {
    const heading = String(values[r][0] || "").trim().toLowerCase();
    if (heading.indexOf("mandatory") >= 0 && heading.indexOf("kpi") >= 0) category = "Mandatory Outcome";
    if (heading.indexOf("critical enabling") >= 0 || heading.indexOf("leading kpi") >= 0) category = "Critical Enabling";
  }

  for (let r = headerRow + 1; r < values.length; r++) {
    const row = values[r];
    const first = String(row[0] || "").trim();
    const firstLower = first.toLowerCase();

    if (STOP_HEADINGS.some(h => firstLower.indexOf(h) === 0)) break;
    if (firstLower.indexOf("mandatory") >= 0 && firstLower.indexOf("kpi") >= 0) {
      category = "Mandatory Outcome";
      continue;
    }
    if (firstLower.indexOf("critical enabling") >= 0 || firstLower.indexOf("leading kpi") >= 0) {
      category = "Critical Enabling";
      continue;
    }
    if (firstLower === "kpi name" || first.indexOf("#REF") === 0) continue;

    const rowType = String(row[7] || "").trim().toLowerCase();
    if (rowType !== "target" && rowType !== "actual") continue;

    if (rowType === "target") {
      const hasName = first.length > 0;
      const definition = String(row[1] || "").trim();
      // A leftover template row: no KPI name and no definition.
      if (!hasName && !definition) {
        current = null;
        continue;
      }
      if (hasName) groupIndex++;
      const next: KPIBlock = {
        kpiName: hasName ? first : (current ? current.kpiName : ""),
        definition: definition,
        operator: String(row[2] || "").trim(),
        target2026: sourceCell(row[3], formats[r][3]),
        unit: String(row[4] || "").trim(),
        nature: String(row[5] || "").trim(),
        frequency: String(row[6] || "").trim(),
        category: category,
        isSubKpi: !hasName,
        groupIndex: groupIndex,
        targetValues: readMonthValues(row, formats[r], monthCols),
        actualValues: readMonthValues([], [], monthCols)
      };
      if (!next.kpiName) {
        current = null;
        continue;
      }
      // A sub-KPI inherits the settings its own row leaves empty.
      const parent = blocks.length > 0 ? blocks[blocks.length - 1] : null;
      if (next.isSubKpi && parent) {
        if (!next.operator) next.operator = parent.operator;
        if (next.target2026.value === "") next.target2026 = parent.target2026;
        if (!next.unit) next.unit = parent.unit;
        if (!next.nature) next.nature = parent.nature;
        if (!next.frequency) next.frequency = parent.frequency;
      }
      if (!next.operator) next.operator = ">=";
      blocks.push(next);
      current = next;
    } else if (current) {
      current.actualValues = readMonthValues(row, formats[r], monthCols);
    }
  }

  // Drop empty template rows: no definition and not a single value.
  return blocks.filter(block =>
    block.definition.length > 0 || block.targetValues.concat(block.actualValues).some(cell => cell.value !== ""));
}

/**
 * Give each sub-KPI of a group its own name, e.g. "Complaint Rate of NPI
 * products" with definitions "Complaint rate for Z10/Z30/Z90" becomes
 * "Complaint Rate of NPI products Z10", "... Z30" and "... Z90".
 */
function nameSubKpis(blocks: KPIBlock[]): void {
  const byGroup: { [groupIndex: number]: KPIBlock[] } = {};
  for (const block of blocks) {
    if (!byGroup[block.groupIndex]) byGroup[block.groupIndex] = [];
    byGroup[block.groupIndex].push(block);
  }
  for (const key of Object.keys(byGroup)) {
    const group = byGroup[Number(key)];
    if (group.length < 2) continue;
    for (const block of group) {
      const shared: { [word: string]: boolean } = {};
      for (const sibling of group) {
        if (sibling === block) continue;
        for (const word of sibling.definition.toLowerCase().split(/\s+/)) shared[word] = true;
      }
      const distinctive = block.definition.split(/\s+/)
        .filter(word => word.length > 0 && !shared[word.toLowerCase()])
        .slice(0, 3)
        .join(" ");
      if (distinctive) block.kpiName = `${block.kpiName} ${distinctive}`;
    }
  }
}

/**
 * Jan..Dec column indexes. The review sheets hold real dates in the header
 * row (Excel serial numbers); text headers such as "Jan-2026" also work.
 */
function functionalMonthColumns(headerRow: Cell[]): { [monthIndex: number]: number } {
  const mapping: { [monthIndex: number]: number } = {};
  for (let c = 0; c < headerRow.length; c++) {
    const raw = headerRow[c];
    if (typeof raw === "number" && raw > 40000 && raw < 60000) {
      const asDate = new Date(Date.UTC(1899, 11, 30) + Math.round(raw) * 86400000);
      const m = asDate.getUTCMonth();
      if (mapping[m] === undefined) mapping[m] = c;
      continue;
    }
    const text = String(raw || "").trim().toLowerCase();
    if (!text) continue;
    for (let m = 0; m < 12; m++) {
      if (text.indexOf(MONTH_PREFIXES[m]) === 0 && mapping[m] === undefined) {
        mapping[m] = c;
        break;
      }
    }
  }
  // Fallback: the calendar columns of a review sheet are I..T.
  for (let m = 0; m < 12; m++) {
    if (mapping[m] === undefined) mapping[m] = 8 + m;
  }
  return mapping;
}

function cleanFunctionName(sheetName: string): string {
  const clean = sheetName.replace(/^[0-9]+[.\s]*/, "").replace(/\s*dmb\s*$/i, "").trim();
  const lower = clean.toLowerCase();
  if (lower.indexOf("customer") >= 0) return "Customer Service";
  if (lower.indexOf("quality") >= 0) return "Quality";
  if (lower.indexOf("regulatory") >= 0) return "Regulatory";
  if (lower.indexOf("isc") >= 0) return "ISC";
  if (lower.indexOf("procurement") >= 0) return "Procurement";
  if (lower.indexOf("marketing") >= 0) return "Marketing";
  if (lower.indexOf("r&d") >= 0) return "R&D";
  if (lower.indexOf("commercial") >= 0) return "Commercial Excellence";
  if (lower.indexOf("nar") >= 0) return "NAR";
  if (lower.indexOf("europe") >= 0) return "Europe";
  if (lower.indexOf("growth") >= 0) return "Growth";
  if (lower.indexOf("finance") >= 0) return "Finance";
  return clean;
}

// =====================================================================
// STRATEGIC EXECUTION DASHBOARD ('AOP Critical')
// =====================================================================

function extractStrategic(sheet: ExcelScript.Worksheet): KPIRow[] {
  const rows: KPIRow[] = [];
  const used = sheet.getUsedRange(true);
  if (!used) return rows;
  const range = sheet.getRangeByIndexes(0, 0,
    used.getRowIndex() + used.getRowCount(),
    Math.max(23, used.getColumnIndex() + used.getColumnCount()));
  const values = range.getValues() as Cell[][];
  const formats = range.getNumberFormats() as string[][];

  let headerRow = 4;
  for (let r = 0; r < Math.min(10, values.length); r++) {
    const text = values[r].map(v => String(v).toLowerCase()).join(" ");
    if (text.indexOf("core kpi") >= 0 || text.indexOf("metric owner") >= 0) {
      headerRow = r;
      break;
    }
  }
  const monthCols = strategicMonthColumns(values[headerRow]);

  let imperative = "Patient Safety & Quality";
  let kpi = "";
  let owner = "";
  let definition = "";
  let aop: SourceCell = NO_CELL;
  let column1 = "";
  let column2 = "";
  let nature = "Higher the better";
  let unit = "%";
  let frequency = "Monthly";

  for (let r = headerRow + 1; r < values.length; r++) {
    const row = values[r];
    const imperativeCell = String(row[0] || "").trim();
    if (imperativeCell && imperativeCell.indexOf("MoS") !== 0 && imperativeCell.indexOf("Strategic") !== 0 &&
        imperativeCell.indexOf("#REF") !== 0) {
      imperative = normalizeImperativeName(imperativeCell);
    }

    const type = String(row[10] || "").trim().toUpperCase();
    const kpiCell = String(row[1] || "").trim();
    const isHeading = kpiCell.toLowerCase().indexOf("value lever") === 0 || kpiCell.indexOf("#REF") === 0;
    if (kpiCell && !isHeading) kpi = kpiCell;

    /*
     * Metadata sits on the Target row. A Target row can leave the KPI name
     * empty and still be a KPI of its own: the (CS) row below "Order Intake
     * Growth" carries the definition "OIT Growth % (CS)".
     */
    const isTarget = type === "T" || type === "TARGET";
    if (isTarget && !isHeading) {
      if (String(row[2] || "").trim()) owner = String(row[2]).trim();
      if (String(row[3] || "").trim()) definition = String(row[3]).trim();
      if (row[4] !== undefined && row[4] !== null && String(row[4]).trim() !== "") aop = sourceCell(row[4], formats[r][4]);
      column1 = String(row[5] || "").trim();
      column2 = String(row[6] || "").trim();
      if (String(row[7] || "").trim()) nature = String(row[7]).trim();
      if (String(row[8] || "").trim()) unit = String(row[8]).trim();
      if (String(row[9] || "").trim()) frequency = String(row[9]).trim();
    }

    if ((isTarget || type === "A" || type === "ACTUAL") && kpi) {
      rows.push(withMonths({
        strategicImperative: imperative,
        sourceSheet: sheet.getName(),
        kpiName: kpi,
        owner: owner,
        definition: definition,
        target2026: aop.value,
        target2026Format: aop.format,
        formats: [],
        column1: column1,
        column2: column2,
        nature: nature,
        unit: unit,
        frequency: frequency,
        rowType: isTarget ? "T" : "A",
        jan: "", feb: "", mar: "", apr: "", may: "", jun: "",
        jul: "", aug: "", sep: "", oct: "", nov: "", dec: "",
        dataType: dataTypeForUnit(unit)
      }, readMonthValues(row, formats[r], monthCols)));
    }
  }
  return rows;
}

/** Jan..Dec columns of 'AOP Critical'; the quarter columns are headed Q1/Q2. */
function strategicMonthColumns(headerRow: Cell[]): { [monthIndex: number]: number } {
  const mapping: { [monthIndex: number]: number } = {};
  for (let m = 0; m < 12; m++) {
    for (let c = 0; c < headerRow.length; c++) {
      const text = String(headerRow[c] || "").trim().toLowerCase();
      if (text.indexOf(MONTH_PREFIXES[m]) === 0 ||
          (m === 2 && text.indexOf("q1") >= 0) || (m === 5 && text.indexOf("q2") >= 0) ||
          (m === 8 && text.indexOf("q3") >= 0) || (m === 11 && text.indexOf("q4") >= 0)) {
        mapping[m] = c;
        break;
      }
    }
  }
  // Fallback: the months of 'AOP Critical' are columns L..W.
  for (let m = 0; m < 12; m++) {
    if (mapping[m] === undefined) mapping[m] = 11 + m;
  }
  return mapping;
}

function normalizeImperativeName(raw: string): string {
  const lower = raw.toLowerCase();
  // "Cusotomer Focus" is spelled that way in the AOP Critical sheet.
  if (lower.indexOf("customer") >= 0 || lower.indexOf("cusotomer") >= 0) return "Customer Focus";
  if (lower.indexOf("deliverability") >= 0 || lower.indexOf("profitability") >= 0) return "Improve Deliverability & Profitability";
  if (lower.indexOf("safety") >= 0 || lower.indexOf("quality") >= 0) return "Patient Safety & Quality";
  if (lower.indexOf("commercial") >= 0 || lower.indexOf("growth") >= 0) return "Drive growth through Commercial Excellence";
  if (lower.indexOf("roadmap") >= 0 || lower.indexOf("innovation") >= 0) return "Roadmap competitiveness & Innovation agility";
  if (lower.indexOf("esg") >= 0 || lower.indexOf("environmental") >= 0 || lower.indexOf("governance") >= 0) {
    return "Environmental, Social & Governance";
  }
  return raw.trim();
}
