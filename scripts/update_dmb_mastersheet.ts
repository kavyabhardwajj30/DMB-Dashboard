/**
 * DMB Dashboard Masterfile Consolidator — Excel Office Script (TypeScript)
 * =========================================================================
 * Purpose:
 *   Extracts KPI Target/Actual values from Functional DMB Review Sheets
 *   and Strategic Execution Dashboard, and populates the canonical Masterfile_DMB_Dashboard.xlsx
 *   with exact Mastersheet layout, colors, banners, and styling.
 */

interface KPIRowData {
  functionName: string;
  sourceSheet: string;
  kpiName: string;
  owner: string;
  definition: string;
  operator?: string;
  target2026: number | string;
  column1?: string;
  column2?: string;
  nature: string;
  unit: string;
  frequency: string;
  rowType: "T" | "A" | "Target" | "Actual";
  monthlyValues: (number | string)[]; // 12 elements (Jan to Dec)
  dataType?: string;
}

/**
 * Main Office Script entry point executed by Power Automate.
 */
function main(workbook: ExcelScript.Workbook): { success: boolean; rowsProcessed: number; message: string } {
  console.log("Starting DMB Masterfile consolidation script...");

  // 1. Identify or Create 'DMB Masterfile' Sheet
  let dmbMasterSheet = workbook.getWorksheet("DMB Masterfile");
  if (!dmbMasterSheet) {
    dmbMasterSheet = workbook.addWorksheet("DMB Masterfile");
  }

  // 2. Identify or Create 'MPR Masterfile' Sheet
  let mprMasterSheet = workbook.getWorksheet("MPR Masterfile");
  if (!mprMasterSheet) {
    mprMasterSheet = workbook.addWorksheet("MPR Masterfile");
  }

  const worksheets = workbook.getWorksheets();
  const functionalKPIs: KPIRowData[] = [];
  const strategicKPIs: KPIRowData[] = [];

  const functionalSheetKeywords = [
    "quality", "regulatory", "isc", "procurement",
    "r&d", "marketing", "customer service", "commercial excellence",
    "nar", "europe", "growth", "finance"
  ];

  for (const sheet of worksheets) {
    const sheetName = sheet.getName().trim();
    const sheetNameLower = sheetName.toLowerCase();

    // Skip output master sheets
    if (sheetName === "DMB Masterfile" || sheetName === "MPR Masterfile") {
      continue;
    }

    // Check if it's Strategic Execution (AOP Critical or Mastersheet)
    if (sheetNameLower.includes("aop critical") || (sheetNameLower.includes("strategic") && !sheetNameLower.includes("dmb"))) {
      const rows = extractStrategicKPIs(sheet);
      strategicKPIs.push(...rows);
      continue;
    }

    // Check if it's Functional MasterSheet
    if (sheetName === "MasterSheet" || sheetName === "Mastersheet") {
      const rows = extractFromFunctionalMasterSheet(sheet);
      if (rows.length > 0) {
        functionalKPIs.length = 0; // Clear any partial rows to use MasterSheet directly
        functionalKPIs.push(...rows);
        break;
      }
    }

    // Check if it's a Functional Review sheet
    const isFunctional = functionalSheetKeywords.some(keyword => sheetNameLower.includes(keyword));
    if (isFunctional) {
      const rows = extractFunctionalSheetKPIs(sheet);
      functionalKPIs.push(...rows);
    }
  }

  console.log(`Extracted ${functionalKPIs.length} functional rows and ${strategicKPIs.length} strategic rows.`);

  // 3. Write Consolidated Data to DMB Masterfile
  const dmbRowsWritten = writeDMBMasterfile(dmbMasterSheet, functionalKPIs);

  // 4. Write Consolidated Data to MPR Masterfile
  const mprRowsWritten = writeMPRMasterfile(mprMasterSheet, strategicKPIs);

  return {
    success: true,
    rowsProcessed: dmbRowsWritten + mprRowsWritten,
    message: `Consolidated ${dmbRowsWritten} DMB rows and ${mprRowsWritten} MPR rows successfully with exact Mastersheet styling.`
  };
}

/**
 * Extract from MasterSheet directly if available.
 */
function extractFromFunctionalMasterSheet(sheet: ExcelScript.Worksheet): KPIRowData[] {
  const used = sheet.getUsedRange();
  if (!used) return [];
  const values = used.getValues();
  if (values.length < 5) return [];

  const rows: KPIRowData[] = [];
  let currentFunction = "Quality";
  let monthCols: { [m: number]: number } = {};

  for (let r = 0; r < values.length; r++) {
    const row = values[r];
    const firstCell = String(row[0] || "").trim();
    if (firstCell.toLowerCase() === "mastersheet" || !firstCell) continue;

    const nextFirst = r + 1 < values.length ? String(values[r + 1][0] || "").trim().toLowerCase() : "";
    if (nextFirst.includes("kpi") || firstCell.toLowerCase().endsWith("dmb")) {
      currentFunction = cleanFunctionName(firstCell);
      continue;
    }

    if (firstCell.toLowerCase().startsWith("kpi name") || firstCell.toLowerCase().startsWith("mandatory kpi")) {
      monthCols = getMonthColumnIndexes(row);
      continue;
    }

    const typeCell = String(row[7] || "").trim().toUpperCase();
    if (typeCell === "TARGET" || typeCell === "ACTUAL" || typeCell === "T" || typeCell === "A") {
      const kpiName = firstCell;
      if (!kpiName || kpiName.startsWith("#REF")) continue;

      const monthlyValues: (number | string)[] = [];
      for (let m = 0; m < 12; m++) {
        const cIdx = monthCols[m];
        const v = cIdx !== undefined && cIdx < row.length ? row[cIdx] : "";
        monthlyValues.push(v !== null && v !== undefined && String(v) !== "#REF!" ? v : "");
      }

      const unitVal = String(row[4] || "").trim();
      rows.push({
        functionName: currentFunction,
        sourceSheet: "MasterSheet",
        kpiName: kpiName,
        owner: "",
        definition: String(row[1] || "").trim(),
        operator: String(row[2] || ">=").trim(),
        target2026: row[3] !== undefined && row[3] !== null ? row[3] : "",
        unit: unitVal,
        nature: String(row[5] || "").trim(),
        frequency: String(row[6] || "Monthly").trim(),
        rowType: typeCell.startsWith("T") ? "Target" : "Actual",
        monthlyValues: monthlyValues,
        dataType: unitVal === "%" ? "Percentage" : "Whole Num"
      });
    }
  }

  return rows;
}

/**
 * Extract Target and Actual KPI rows from a functional review worksheet.
 */
function extractFunctionalSheetKPIs(sheet: ExcelScript.Worksheet): KPIRowData[] {
  const sheetName = sheet.getName().trim();
  const functionName = cleanFunctionName(sheetName);
  const usedRange = sheet.getUsedRange();
  if (!usedRange) return [];
  const values = usedRange.getValues();
  if (values.length < 5) return [];

  const rows: KPIRowData[] = [];
  let headerRowIndex = 4;
  for (let r = 0; r < Math.min(10, values.length); r++) {
    const rowStr = values[r].map(v => String(v).toLowerCase()).join(" ");
    if (rowStr.includes("jan") || rowStr.includes("target/ actual") || rowStr.includes("target aop")) {
      headerRowIndex = r;
      break;
    }
  }

  const headerRow = values[headerRowIndex];
  const monthCols = getMonthColumnIndexes(headerRow);

  for (let r = headerRowIndex + 1; r < values.length; r++) {
    const row = values[r];
    const typeVal = String(row[7] || "").trim().toUpperCase();
    if (typeVal === "TARGET" || typeVal === "ACTUAL" || typeVal === "T" || typeVal === "A") {
      const kpiName = String(row[0] || "").trim();
      if (!kpiName || kpiName.startsWith("Value Lever") || kpiName.startsWith("#REF")) continue;

      const monthlyValues: (number | string)[] = [];
      for (let m = 0; m < 12; m++) {
        const cIdx = monthCols[m];
        const v = cIdx !== undefined && cIdx < row.length ? row[cIdx] : "";
        monthlyValues.push(v !== null && v !== undefined && String(v) !== "#REF!" ? v : "");
      }

      const unitVal = String(row[4] || "").trim();
      rows.push({
        functionName: functionName,
        sourceSheet: sheetName,
        kpiName: kpiName,
        owner: "",
        definition: String(row[1] || "").trim(),
        operator: String(row[2] || ">=").trim(),
        target2026: row[3] !== undefined && row[3] !== null ? row[3] : "",
        unit: unitVal,
        nature: String(row[5] || "").trim(),
        frequency: String(row[6] || "Monthly").trim(),
        rowType: typeVal.startsWith("T") ? "Target" : "Actual",
        monthlyValues: monthlyValues,
        dataType: unitVal === "%" ? "Percentage" : "Whole Num"
      });
    }
  }

  return rows;
}

/**
 * Extract Strategic Imperative KPIs from the AOP Critical worksheet.
 */
function extractStrategicKPIs(sheet: ExcelScript.Worksheet): KPIRowData[] {
  const usedRange = sheet.getUsedRange();
  if (!usedRange) return [];
  const values = usedRange.getValues();
  if (values.length < 5) return [];

  const rows: KPIRowData[] = [];
  let headerRowIndex = 4;
  for (let r = 0; r < Math.min(12, values.length); r++) {
    const rowStr = values[r].map(v => String(v).toLowerCase()).join(" ");
    if (rowStr.includes("jan") || rowStr.includes("feb")) {
      headerRowIndex = r;
      break;
    }
  }

  const headerRow = values[headerRowIndex];
  const monthCols = getStrategicMonthCols(headerRow);
  let currentImperative = "Customer Focus";

  for (let r = headerRowIndex + 1; r < values.length; r++) {
    const row = values[r];
    const impVal = String(row[0] || "").trim();
    if (impVal && !impVal.startsWith("MoS") && !impVal.startsWith("Strategic")) {
      currentImperative = normalizeImperativeName(impVal);
    }

    const rowTypeVal = String(row[10] || "").trim().toUpperCase();
    if (rowTypeVal === "T" || rowTypeVal === "A") {
      let kpiName = String(row[1] || "").trim();
      if (!kpiName || kpiName.toLowerCase().startsWith("value lever")) continue;

      const kpiDef = String(row[3] || "").trim();
      if (kpiDef.includes("(EQ)") && !kpiName.includes("(EQ)")) {
        kpiName = `${kpiName} (EQ)`;
      } else if (kpiDef.includes("(CS)") && !kpiName.includes("(CS)")) {
        kpiName = `${kpiName} (CS)`;
      }

      const monthlyValues: (number | string)[] = [];
      for (let m = 0; m < 12; m++) {
        const cIdx = monthCols[m];
        const v = cIdx !== undefined && cIdx < row.length ? row[cIdx] : "";
        monthlyValues.push(v !== null && v !== undefined && String(v) !== "#REF!" ? v : "");
      }

      const unitVal = String(row[8] || "").trim();
      rows.push({
        functionName: currentImperative,
        sourceSheet: sheet.getName(),
        kpiName: kpiName,
        owner: String(row[2] || "").trim(),
        definition: kpiDef,
        target2026: row[4] !== undefined ? row[4] : "",
        nature: String(row[7] || "").trim(),
        unit: unitVal,
        frequency: String(row[9] || "").trim(),
        rowType: rowTypeVal as "T" | "A",
        monthlyValues: monthlyValues,
        dataType: unitVal === "%" ? "Percentage" : "Whole Num"
      });
    }
  }

  return rows;
}

/**
 * Write structured rows into DMB Masterfile with exact Mastersheet layout & styling.
 */
function writeDMBMasterfile(sheet: ExcelScript.Worksheet, rows: KPIRowData[]): number {
  if (rows.length === 0) return 0;
  sheet.getUsedRange()?.clear();

  const dmbHeaders = [
    "KPI Name",
    "KPI  Definition",
    "Operators",
    "Target AOP 2026",
    "Units",
    "Metric nature",
    "Frequency",
    "Target/ Actual",
    "Jan-2026",
    "Feb-2026",
    "Mar-2026",
    "Apr-2026",
    "May-2026",
    "Jun-2026",
    "Jul-2026",
    "Aug-2026",
    "Sep-2026",
    "Oct-2026",
    "Nov-2026",
    "Dec-2026",
    "Data Type"
  ];

  const functionGroups = new Map<string, KPIRowData[]>();
  for (const r of rows) {
    const fn = r.functionName || "General";
    if (!functionGroups.has(fn)) {
      functionGroups.set(fn, []);
    }
    functionGroups.get(fn)!.push(r);
  }

  const tableData: (string | number)[][] = [];
  // Row 1: Mastersheet Title Banner
  tableData.push(["Mastersheet", ...Array(dmbHeaders.length - 1).fill("")]);
  // Row 2: Blank
  tableData.push(Array(dmbHeaders.length).fill(""));

  const sectionHeaderRows: number[] = [];
  const colHeaderRows: number[] = [];
  const targetRows: number[] = [];
  const percentRows: number[] = [];

  for (const [funcName, kpis] of functionGroups) {
    const secRowIdx = tableData.length;
    sectionHeaderRows.push(secRowIdx);
    const displayName = funcName.toLowerCase().endsWith("dmb") ? funcName : `${funcName} DMB`;
    tableData.push([displayName, ...Array(dmbHeaders.length - 1).fill("")]);

    const colRowIdx = tableData.length;
    colHeaderRows.push(colRowIdx);
    tableData.push([...dmbHeaders]);

    const kpiMap = new Map<string, { target?: KPIRowData; actual?: KPIRowData }>();
    for (const item of kpis) {
      if (!kpiMap.has(item.kpiName)) {
        kpiMap.set(item.kpiName, {});
      }
      const pair = kpiMap.get(item.kpiName)!;
      const tUpper = String(item.rowType).toUpperCase();
      if (tUpper.startsWith("T")) pair.target = item;
      else pair.actual = item;
    }

    for (const [kpiName, pair] of kpiMap) {
      const base = pair.target || pair.actual!;
      const def = base.definition || "";
      const op = base.operator || ">=";
      const aop = base.target2026 !== undefined ? base.target2026 : "";
      const unit = base.unit || "%";
      const nature = base.nature || "Higher the better";
      const freq = base.frequency || "Monthly";
      const dType = base.dataType || (unit === "%" ? "Percentage" : "Whole Num");

      // Target row
      const targetRowIdx = tableData.length;
      targetRows.push(targetRowIdx);
      if (unit === "%") percentRows.push(targetRowIdx);

      const tVals = pair.target ? pair.target.monthlyValues : Array(12).fill("");
      tableData.push([
        kpiName,
        def,
        op,
        aop,
        unit,
        nature,
        freq,
        "Target",
        ...tVals,
        dType
      ]);

      // Actual row
      const actualRowIdx = tableData.length;
      if (unit === "%") percentRows.push(actualRowIdx);

      const aVals = pair.actual ? pair.actual.monthlyValues : Array(12).fill("");
      tableData.push([
        kpiName,
        def,
        op,
        aop,
        unit,
        nature,
        freq,
        "Actual",
        ...aVals,
        dType
      ]);
    }

    tableData.push(Array(dmbHeaders.length).fill(""));
  }

  const range = sheet.getRangeByIndexes(0, 0, tableData.length, dmbHeaders.length);
  range.setValues(tableData);

  // 1. Style Top Title Banner (Row 1)
  const topTitle = sheet.getRangeByIndexes(0, 0, 1, dmbHeaders.length);
  topTitle.getFormat().getFill().setColor("#002D4B");
  topTitle.getFormat().getFont().setColor("#FFFFFF");
  topTitle.getFormat().getFont().setBold(true);
  topTitle.getFormat().getFont().setSize(12);

  // 2. Style Section Headers
  for (const rIdx of sectionHeaderRows) {
    const secRange = sheet.getRangeByIndexes(rIdx, 0, 1, 1);
    secRange.getFormat().getFont().setBold(true);
    secRange.getFormat().getFont().setSize(11);
    secRange.getFormat().getFont().setColor("#000000");
  }

  // 3. Style Column Headers (Dark Blue with White Text)
  for (const rIdx of colHeaderRows) {
    const headerRange = sheet.getRangeByIndexes(rIdx, 0, 1, dmbHeaders.length);
    headerRange.getFormat().getFill().setColor("#004071");
    headerRange.getFormat().getFont().setColor("#FFFFFF");
    headerRange.getFormat().getFont().setBold(true);
  }

  // 4. Style Target Rows (Light Blue #D9EEF8 across the entire row)
  for (const rIdx of targetRows) {
    const tRange = sheet.getRangeByIndexes(rIdx, 0, 1, dmbHeaders.length);
    tRange.getFormat().getFill().setColor("#D9EEF8");
  }

  // 5. Format Percentage values
  for (const rIdx of percentRows) {
    const valRange = sheet.getRangeByIndexes(rIdx, 8, 1, 12);
    valRange.setNumberFormatLocal("0.0%");
  }

  sheet.getUsedRange()?.getFormat().autofitColumns();
  return tableData.length;
}

/**
 * Write structured rows into MPR Masterfile with exact Mastersheet layout & styling.
 */
function writeMPRMasterfile(sheet: ExcelScript.Worksheet, rows: KPIRowData[]): number {
  if (rows.length === 0) return 0;
  sheet.getUsedRange()?.clear();

  const mprHeaders = [
    "Core KPI's",
    "Metric Owner",
    "KPI  Definition",
    "AOP 2026",
    "Column1",
    "Column2",
    "Metric nature",
    "Units",
    "Frequency",
    "Target/\nActual",
    "Jan",
    "Feb",
    "Q1'26",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec"
  ];

  const imperativeGroups = new Map<string, KPIRowData[]>();
  for (const r of rows) {
    const imp = r.functionName || "Strategic Imperative";
    if (!imperativeGroups.has(imp)) {
      imperativeGroups.set(imp, []);
    }
    imperativeGroups.get(imp)!.push(r);
  }

  const tableData: (string | number)[][] = [];
  // Row 1: Title Banner
  tableData.push(["Strategic Imperatives", ...Array(mprHeaders.length - 1).fill("")]);
  // Row 2: Blank
  tableData.push(Array(mprHeaders.length).fill(""));

  const sectionHeaderRows: number[] = [];
  const colHeaderRows: number[] = [];
  const targetRows: number[] = [];
  const percentRows: number[] = [];

  for (const [imperative, kpis] of imperativeGroups) {
    const secRowIdx = tableData.length;
    sectionHeaderRows.push(secRowIdx);
    tableData.push([imperative, ...Array(mprHeaders.length - 1).fill("")]);

    const colRowIdx = tableData.length;
    colHeaderRows.push(colRowIdx);
    tableData.push([...mprHeaders]);

    const kpiMap = new Map<string, { target?: KPIRowData; actual?: KPIRowData }>();
    for (const item of kpis) {
      if (!kpiMap.has(item.kpiName)) {
        kpiMap.set(item.kpiName, {});
      }
      const pair = kpiMap.get(item.kpiName)!;
      const tUpper = String(item.rowType).toUpperCase();
      if (tUpper.startsWith("T")) pair.target = item;
      else pair.actual = item;
    }

    for (const [kpiName, pair] of kpiMap) {
      const base = pair.target || pair.actual!;
      const owner = base.owner || "";
      const def = base.definition || "";
      const aop = base.target2026 !== undefined ? base.target2026 : "";
      const nature = base.nature || "Higher the better";
      const unit = base.unit || "%";
      const freq = base.frequency || "Monthly";

      // Target row (T)
      const targetRowIdx = tableData.length;
      targetRows.push(targetRowIdx);
      if (unit === "%") percentRows.push(targetRowIdx);

      const tVals = pair.target ? pair.target.monthlyValues : Array(12).fill("");
      tableData.push([
        kpiName,
        owner,
        def,
        aop,
        "",
        "",
        nature,
        unit,
        freq,
        "T",
        ...tVals
      ]);

      // Actual row (A)
      const actualRowIdx = tableData.length;
      if (unit === "%") percentRows.push(actualRowIdx);

      const aVals = pair.actual ? pair.actual.monthlyValues : Array(12).fill("");
      tableData.push([
        kpiName,
        owner,
        def,
        aop,
        "",
        "",
        nature,
        unit,
        freq,
        "A",
        ...aVals
      ]);
    }

    tableData.push(Array(mprHeaders.length).fill(""));
  }

  const range = sheet.getRangeByIndexes(0, 0, tableData.length, mprHeaders.length);
  range.setValues(tableData);

  // 1. Style Top Title Banner (Row 1)
  const topTitle = sheet.getRangeByIndexes(0, 0, 1, mprHeaders.length);
  topTitle.getFormat().getFill().setColor("#002D4B");
  topTitle.getFormat().getFont().setColor("#FFFFFF");
  topTitle.getFormat().getFont().setBold(true);
  topTitle.getFormat().getFont().setSize(12);

  // 2. Style Section Headers
  for (const rIdx of sectionHeaderRows) {
    const secRange = sheet.getRangeByIndexes(rIdx, 0, 1, 1);
    secRange.getFormat().getFont().setBold(true);
    secRange.getFormat().getFont().setSize(11);
    secRange.getFormat().getFont().setColor("#000000");
  }

  // 3. Style Column Headers (Dark Blue with White Text)
  for (const rIdx of colHeaderRows) {
    const headerRange = sheet.getRangeByIndexes(rIdx, 0, 1, mprHeaders.length);
    headerRange.getFormat().getFill().setColor("#004071");
    headerRange.getFormat().getFont().setColor("#FFFFFF");
    headerRange.getFormat().getFont().setBold(true);
  }

  // 4. Style Target Rows (Light Blue #D9EEF8 across the entire row)
  for (const rIdx of targetRows) {
    const tRange = sheet.getRangeByIndexes(rIdx, 0, 1, mprHeaders.length);
    tRange.getFormat().getFill().setColor("#D9EEF8");
  }

  // 5. Format Percentage values
  for (const rIdx of percentRows) {
    const valRange = sheet.getRangeByIndexes(rIdx, 10, 1, 12);
    valRange.setNumberFormatLocal("0.0%");
  }

  sheet.getUsedRange()?.getFormat().autofitColumns();
  return tableData.length;
}

function getMonthColumnIndexes(headerRow: (string | number | boolean)[]): { [monthIndex: number]: number } {
  const prefixes = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const mapping: { [monthIndex: number]: number } = {};
  for (let m = 0; m < prefixes.length; m++) {
    const p = prefixes[m];
    for (let c = 0; c < headerRow.length; c++) {
      const val = String(headerRow[c] || "").trim().toLowerCase();
      if (val.startsWith(p) || val.includes(`-${p}`) || val.includes(p)) {
        mapping[m] = c;
        break;
      }
    }
  }
  for (let m = 0; m < 12; m++) {
    if (mapping[m] === undefined) {
      mapping[m] = 8 + m;
    }
  }
  return mapping;
}

function getStrategicMonthCols(headerRow: (string | number | boolean)[]): { [monthIndex: number]: number } {
  const prefixes = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const mapping: { [monthIndex: number]: number } = {};
  for (let m = 0; m < prefixes.length; m++) {
    const p = prefixes[m];
    for (let c = 0; c < headerRow.length; c++) {
      const val = String(headerRow[c] || "").trim().toLowerCase();
      if (val.startsWith(p) || (m === 2 && val.includes("q1")) || (m === 5 && val.includes("q2")) || (m === 8 && val.includes("q3")) || (m === 11 && val.includes("q4"))) {
        mapping[m] = c;
        break;
      }
    }
  }
  for (let m = 0; m < 12; m++) {
    if (mapping[m] === undefined) {
      mapping[m] = 10 + m;
    }
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

function normalizeImperativeName(raw: string): string {
  const lower = raw.toLowerCase();
  if (lower.includes("customer")) return "Customer Focus";
  if (lower.includes("deliverability") || lower.includes("profitability")) return "Improve Deliverability & Profitability";
  if (lower.includes("safety") || lower.includes("quality")) return "Patient Safety & Quality";
  if (lower.includes("commercial") || lower.includes("growth")) return "Drive growth through Commercial Excellence";
  if (lower.includes("roadmap") || lower.includes("innovation")) return "Roadmap competitiveness & Innovation agility";
  if (lower.includes("esg") || lower.includes("environmental") || lower.includes("governance")) return "Environmental, Social & Governance (ESG)";
  return raw.trim();
}
