/**
 * SCRIPT 1: Extract Functional DMB Review Data (Office Script / TypeScript)
 * =========================================================================
 * Location: Run on 'Functional DMB Review Sheets-17th_sept.xlsx' in SharePoint / Excel Online.
 * Action in Power Automate: "Run script" -> Select Functional DMB workbook.
 * Returns: JSON string containing extracted functional KPI rows with exact metadata & values.
 *
 * Reads only the visible function tabs (1.Quality ... 12.FINANCE) and only the
 * KPI table at the top of each tab, stopping at the Trends / Root Cause
 * Analysis / Action tracker blocks below it. The hidden 'MasterSheet',
 * 'MPR Review', 'Marketing June' and 'Sheet1 (2)' tabs are skipped, so old and
 * broken (#REF!) copies never reach the Masterfile.
 */

interface FunctionalKPIRow {
  functionName: string;
  sourceSheet: string;
  kpiName: string;
  definition: string;
  operator: string;
  target2026: number | string;
  unit: string;
  nature: string;
  frequency: string;
  rowType: "Target" | "Actual" | "T" | "A";
  jan: number | string;
  feb: number | string;
  mar: number | string;
  apr: number | string;
  may: number | string;
  jun: number | string;
  jul: number | string;
  aug: number | string;
  sep: number | string;
  oct: number | string;
  nov: number | string;
  dec: number | string;
  dataType?: string;
  category?: string;
}

/** A KPI and the Target / Actual values read from one function tab. */
interface KPIBlock {
  kpiName: string;
  definition: string;
  operator: string;
  target2026: number | string;
  unit: string;
  nature: string;
  frequency: string;
  category: string;
  isSubKpi: boolean;
  groupIndex: number;
  targetValues: (number | string)[];
  actualValues: (number | string)[];
}

/** Tabs that never hold current KPI data. */
const SKIP_SHEETS = ["mpr review", "mastersheet", "sheet1 (2)", "marketing june"];

/** Everything below these headings is RCA / action-tracker content. */
const STOP_HEADINGS = [
  "trends", "root cause analysis", "paretos", "action tracker",
  "if the kpi is reported", "handshake"
];

function main(workbook: ExcelScript.Workbook): string {
  console.log("Starting functional data extraction...");
  const extractedRows: FunctionalKPIRow[] = [];

  const worksheets = workbook.getWorksheets();
  const functionalKeywords = [
    "quality", "regulatory", "isc", "procurement",
    "r&d", "marketing", "customer service", "commercial excellence",
    "nar", "europe", "growth", "finance"
  ];

  for (const sheet of worksheets) {
    const sName = sheet.getName().trim();
    const sNameLower = sName.toLowerCase();

    // Only visible function tabs; hidden copies hold stale or #REF! data.
    if (sheet.getVisibility() !== ExcelScript.SheetVisibility.visible) continue;
    if (SKIP_SHEETS.indexOf(sNameLower) >= 0) continue;
    if (!functionalKeywords.some(k => sNameLower.includes(k))) continue;

    const used = sheet.getUsedRange();
    if (!used) continue;
    const values = used.getValues();
    if (values.length < 5) continue;

    const functionName = cleanFunctionName(sName);

    // Locate header row
    let headerRowIdx = 4;
    for (let r = 0; r < Math.min(10, values.length); r++) {
      const firstCell = String(values[r][0] || "").trim().toLowerCase();
      if (firstCell === "kpi name" || firstCell.indexOf("mandatory kpi") === 0) {
        headerRowIdx = r;
        break;
      }
    }

    const monthCols = getMonthColumnIndexes(values[headerRowIdx]);
    const blocks = readKPIBlocks(values, headerRowIdx, monthCols);
    nameSubKpis(blocks);

    for (const block of blocks) {
      let dType = "Whole Number";
      if (block.unit === "%") dType = "Percentage";
      else if (block.unit === "Mn" || block.unit === "K" || block.unit.toLowerCase().includes("decimal")) dType = "Decimal";

      for (const rowType of ["Target", "Actual"]) {
        const monthly = rowType === "Target" ? block.targetValues : block.actualValues;
        extractedRows.push({
          functionName: functionName,
          sourceSheet: sName,
          kpiName: block.kpiName,
          definition: block.definition,
          operator: block.operator,
          target2026: block.target2026,
          unit: block.unit,
          nature: block.nature,
          frequency: block.frequency,
          rowType: rowType === "Target" ? "Target" : "Actual",
          jan: monthly[0], feb: monthly[1], mar: monthly[2], apr: monthly[3],
          may: monthly[4], jun: monthly[5], jul: monthly[6], aug: monthly[7],
          sep: monthly[8], oct: monthly[9], nov: monthly[10], dec: monthly[11],
          dataType: dType,
          category: block.category
        });
      }
    }
  }

  console.log(`Extracted ${extractedRows.length} functional KPI rows.`);
  return JSON.stringify(extractedRows);
}

/**
 * Read the KPI table of one function tab.
 * A KPI row carries its name in column A; a sub-KPI (for example the Z30 and
 * Z90 complaint rates) leaves column A empty and names itself in the KPI
 * definition, so it is kept with the KPI name of the row above it.
 */
function readKPIBlocks(
  values: (string | number | boolean)[][],
  headerRowIdx: number,
  monthCols: { [monthIndex: number]: number }
): KPIBlock[] {
  const blocks: KPIBlock[] = [];
  let category = "";
  let groupIndex = -1;
  let current: KPIBlock | null = null;

  for (let r = headerRowIdx + 1; r < values.length; r++) {
    const row = values[r];
    const firstCell = String(row[0] || "").trim();
    const firstLower = firstCell.toLowerCase();

    if (STOP_HEADINGS.some(h => firstLower.indexOf(h) === 0)) break;

    if (firstLower.indexOf("mandatory") >= 0 && firstLower.indexOf("kpi") >= 0) {
      category = "Mandatory Outcome";
      continue;
    }
    if (firstLower.indexOf("critical enabling") >= 0 || firstLower.indexOf("leading kpi") >= 0) {
      category = "Critical Enabling";
      continue;
    }
    if (firstLower === "kpi name" || firstCell.indexOf("#REF") === 0) continue;

    const rowType = String(row[7] || "").trim().toLowerCase();
    if (rowType !== "target" && rowType !== "actual") continue;

    const monthly = readMonthValues(row, monthCols);

    if (rowType === "target") {
      const hasName = firstCell.length > 0;
      const definition = String(row[1] || "").trim();

      // A leftover template row: no KPI name and no definition. A real sub-KPI
      // always names itself in the definition, e.g. "Complaint rate for Z30".
      if (!hasName && !definition) {
        current = null;
        continue;
      }

      if (hasName) groupIndex++;
      current = {
        kpiName: hasName ? firstCell : (current ? current.kpiName : ""),
        definition: String(row[1] || "").trim(),
        operator: String(row[2] || ">=").trim(),
        target2026: row[3] !== undefined && row[3] !== null ? row[3] : "",
        unit: String(row[4] || "").trim(),
        nature: String(row[5] || "").trim(),
        frequency: String(row[6] || "").trim(),
        category: category,
        isSubKpi: !hasName,
        groupIndex: groupIndex,
        targetValues: monthly,
        actualValues: ["", "", "", "", "", "", "", "", "", "", "", ""]
      };
      if (!current.kpiName) {
        current = null;
        continue;
      }
      // A sub-KPI inherits the settings its own row leaves empty.
      const parent = blocks.length > 0 ? blocks[blocks.length - 1] : null;
      if (current.isSubKpi && parent) {
        if (!current.operator) current.operator = parent.operator;
        if (current.target2026 === "") current.target2026 = parent.target2026;
        if (!current.unit) current.unit = parent.unit;
        if (!current.nature) current.nature = parent.nature;
        if (!current.frequency) current.frequency = parent.frequency;
      }
      blocks.push(current);
    } else if (current) {
      current.actualValues = monthly;
    }
  }

  // Drop empty template rows: no definition and not a single value.
  return blocks.filter(block => {
    const hasValue = block.targetValues.concat(block.actualValues).some(v => v !== "");
    return hasValue || block.definition.length > 0;
  });
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
        for (const word of sibling.definition.toLowerCase().split(/\s+/)) {
          shared[word] = true;
        }
      }
      const distinctive = block.definition.split(/\s+/)
        .filter(word => word.length > 0 && !shared[word.toLowerCase()])
        .slice(0, 3)
        .join(" ");
      if (distinctive) block.kpiName = `${block.kpiName} ${distinctive}`;
    }
  }
}

function readMonthValues(
  row: (string | number | boolean)[],
  monthCols: { [monthIndex: number]: number }
): (number | string)[] {
  const monthly: (number | string)[] = [];
  for (let m = 0; m < 12; m++) {
    const cIdx = monthCols[m];
    let value: number | string = "";
    if (cIdx !== undefined && cIdx < row.length) {
      const raw = row[cIdx];
      const text = String(raw);
      if (raw !== null && raw !== undefined && text !== "" && text !== "#REF!" && text !== "None") {
        value = raw as number | string;
      }
    }
    monthly.push(value);
  }
  return monthly;
}

/**
 * Map Jan..Dec to their column indexes. The review sheets hold real dates in
 * the header row, which arrive as Excel serial numbers, so those are converted
 * back to a month. Text headers such as "Jan-2026" are also supported.
 */
function getMonthColumnIndexes(headerRow: (string | number | boolean)[]): { [monthIndex: number]: number } {
  const prefixes = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const mapping: { [monthIndex: number]: number } = {};

  for (let c = 0; c < headerRow.length; c++) {
    const raw = headerRow[c];
    if (typeof raw === "number" && raw > 40000 && raw < 60000) {
      // Excel serial date -> month index (day 0 is 1899-12-30).
      const asDate = new Date(Date.UTC(1899, 11, 30) + Math.round(raw) * 86400000);
      const monthIndex = asDate.getUTCMonth();
      if (mapping[monthIndex] === undefined) mapping[monthIndex] = c;
      continue;
    }
    const text = String(raw || "").trim().toLowerCase();
    if (!text) continue;
    for (let m = 0; m < prefixes.length; m++) {
      if (text.indexOf(prefixes[m]) === 0 && mapping[m] === undefined) {
        mapping[m] = c;
        break;
      }
    }
  }

  // Fallback: the calendar columns of a review sheet are I..T (8 to 19).
  for (let m = 0; m < 12; m++) {
    if (mapping[m] === undefined) mapping[m] = 8 + m;
  }

  return mapping;
}

function cleanFunctionName(sheetName: string): string {
  const clean = sheetName.replace(/^[0-9]+[.\s]*/, "").replace(/\s*dmb\s*$/i, "").trim();
  const lower = clean.toLowerCase();
  if (lower.includes("customer")) return "Customer Service";
  if (lower.includes("quality")) return "Quality";
  if (lower.includes("regulatory")) return "Regulatory";
  if (lower.includes("isc")) return "ISC";
  if (lower.includes("procurement")) return "Procurement";
  if (lower.includes("marketing")) return "Marketing";
  if (lower.includes("r&d")) return "R&D";
  if (lower.includes("commercial")) return "Commercial Excellence";
  if (lower.includes("nar")) return "NAR";
  if (lower.includes("europe")) return "Europe";
  if (lower.includes("growth")) return "Growth";
  if (lower.includes("finance")) return "Finance";
  return clean;
}
