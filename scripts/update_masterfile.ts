/**
 * SCRIPT 3: Populate Masterfile Worksheets (Office Script / TypeScript)
 * =====================================================================
 * Location: Run on 'Masterfile_DMB_Dashboard.xlsx' in SharePoint / Excel Online.
 * Action in Power Automate: "Run script" -> Select Masterfile workbook.
 * Parameters:
 *   - functionalJson (string): JSON output from SCRIPT 1 (extract_functional_dmb)
 *   - strategicJson (string): JSON output from SCRIPT 2 (extract_strategic_execution)
 */

interface KPIRow {
  functionName?: string;
  strategicImperative?: string;
  sourceSheet?: string;
  kpiName: string;
  owner?: string;
  definition: string;
  operator?: string;
  target2026: number | string;
  column1?: string;
  column2?: string;
  nature: string;
  unit: string;
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
}

function main(workbook: ExcelScript.Workbook, functionalJson?: string, strategicJson?: string): { dmbRows: number; mprRows: number; status: string } {
  console.log("Starting Masterfile sync with exact Mastersheet styling...");

  // Safety check: if run manually in Excel without Power Automate parameters
  const hasFunctionalInput = functionalJson && functionalJson.trim().length > 2;
  const hasStrategicInput = strategicJson && strategicJson.trim().length > 2;

  if (!hasFunctionalInput && !hasStrategicInput) {
    const msg = "Notice: 'Update_Masterfile' requires JSON inputs from Power Automate (Step 1 & Step 2). When clicked manually in Excel, no data is passed. Please run the Power Automate flow to execute the sync.";
    console.warn(msg);
    return {
      dmbRows: 0,
      mprRows: 0,
      status: msg
    };
  }

  let dmbRowsCount = 0;
  let mprRowsCount = 0;

  // 1. Process Functional Review Data into 'DMB Masterfile'
  if (hasFunctionalInput) {
    try {
      const functionalRows: KPIRow[] = JSON.parse(functionalJson!);
      let dmbSheet = workbook.getWorksheet("DMB Masterfile");
      if (!dmbSheet) {
        dmbSheet = workbook.addWorksheet("DMB Masterfile");
      }
      dmbRowsCount = writeDMBMasterfileExact(dmbSheet, functionalRows);
      console.log(`Successfully formatted and wrote ${dmbRowsCount} rows to DMB Masterfile.`);
    } catch (err) {
      console.error("Error writing DMB Masterfile:", err);
    }
  }

  // 2. Process Strategic Execution Data into 'MPR Masterfile'
  if (hasStrategicInput) {
    try {
      const strategicRows: KPIRow[] = JSON.parse(strategicJson!);
      let mprSheet = workbook.getWorksheet("MPR Masterfile");
      if (!mprSheet) {
        mprSheet = workbook.addWorksheet("MPR Masterfile");
      }
      mprRowsCount = writeMPRMasterfileExact(mprSheet, strategicRows);
      console.log(`Successfully formatted and wrote ${mprRowsCount} rows to MPR Masterfile.`);
    } catch (err) {
      console.error("Error writing MPR Masterfile:", err);
    }
  }

  return {
    dmbRows: dmbRowsCount,
    mprRows: mprRowsCount,
    status: `Masterfile successfully synced with exact Mastersheet styling (${dmbRowsCount} DMB rows, ${mprRowsCount} MPR rows).`
  };
}

/**
 * Writes Functional review data in the EXACT layout and styling of the Mastersheet.
 */
function writeDMBMasterfileExact(sheet: ExcelScript.Worksheet, rows: KPIRow[]): number {
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

  // Group rows by function preserving natural order
  const functionGroups = new Map<string, KPIRow[]>();
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
    // Section Header (e.g. 'Quality DMB')
    const secRowIdx = tableData.length;
    sectionHeaderRows.push(secRowIdx);
    const displayName = funcName.toLowerCase().endsWith("dmb") ? funcName : `${funcName} DMB`;
    tableData.push([displayName, ...Array(dmbHeaders.length - 1).fill("")]);

    // Column Headers Row
    const colRowIdx = tableData.length;
    colHeaderRows.push(colRowIdx);
    tableData.push([...dmbHeaders]);

    // Pair Target and Actual rows per KPI
    const kpiMap = new Map<string, { target?: KPIRow; actual?: KPIRow }>();
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

      const t = pair.target;
      tableData.push([
        kpiName,
        def,
        op,
        aop,
        unit,
        nature,
        freq,
        "Target",
        t ? t.jan : "",
        t ? t.feb : "",
        t ? t.mar : "",
        t ? t.apr : "",
        t ? t.may : "",
        t ? t.jun : "",
        t ? t.jul : "",
        t ? t.aug : "",
        t ? t.sep : "",
        t ? t.oct : "",
        t ? t.nov : "",
        t ? t.dec : "",
        dType
      ]);

      // Actual row
      const actualRowIdx = tableData.length;
      if (unit === "%") percentRows.push(actualRowIdx);

      const a = pair.actual;
      tableData.push([
        kpiName,
        def,
        op,
        aop,
        unit,
        nature,
        freq,
        "Actual",
        a ? a.jan : "",
        a ? a.feb : "",
        a ? a.mar : "",
        a ? a.apr : "",
        a ? a.may : "",
        a ? a.jun : "",
        a ? a.jul : "",
        a ? a.aug : "",
        a ? a.sep : "",
        a ? a.oct : "",
        a ? a.nov : "",
        a ? a.dec : "",
        dType
      ]);
    }

    // Blank separator row
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

  // 2. Style Section Headers (Quality DMB, Regulatory, etc.)
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

  // Auto-fit columns
  sheet.getUsedRange()?.getFormat().autofitColumns();

  return tableData.length;
}

/**
 * Writes Strategic execution data in the EXACT layout and styling of the Mastersheet.
 */
function writeMPRMasterfileExact(sheet: ExcelScript.Worksheet, rows: KPIRow[]): number {
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

  const imperativeGroups = new Map<string, KPIRow[]>();
  for (const r of rows) {
    const imp = r.strategicImperative || "Strategic Imperative";
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
    // Section Header (e.g. 'Patient Safety & Quality')
    const secRowIdx = tableData.length;
    sectionHeaderRows.push(secRowIdx);
    tableData.push([imperative, ...Array(mprHeaders.length - 1).fill("")]);

    // Column Headers Row
    const colRowIdx = tableData.length;
    colHeaderRows.push(colRowIdx);
    tableData.push([...mprHeaders]);

    // Pair Target (T) and Actual (A)
    const kpiMap = new Map<string, { target?: KPIRow; actual?: KPIRow }>();
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
      const col1 = base.column1 || "";
      const col2 = base.column2 || "";
      const nature = base.nature || "Higher the better";
      const unit = base.unit || "%";
      const freq = base.frequency || "Monthly";

      // Target row (T)
      const targetRowIdx = tableData.length;
      targetRows.push(targetRowIdx);
      if (unit === "%") percentRows.push(targetRowIdx);

      const t = pair.target;
      tableData.push([
        kpiName,
        owner,
        def,
        aop,
        col1,
        col2,
        nature,
        unit,
        freq,
        "T",
        t ? t.jan : "",
        t ? t.feb : "",
        t ? t.mar : "",
        t ? t.apr : "",
        t ? t.may : "",
        t ? t.jun : "",
        t ? t.jul : "",
        t ? t.aug : "",
        t ? t.sep : "",
        t ? t.oct : "",
        t ? t.nov : "",
        t ? t.dec : ""
      ]);

      // Actual row (A)
      const actualRowIdx = tableData.length;
      if (unit === "%") percentRows.push(actualRowIdx);

      const a = pair.actual;
      tableData.push([
        kpiName,
        owner,
        def,
        aop,
        col1,
        col2,
        nature,
        unit,
        freq,
        "A",
        a ? a.jan : "",
        a ? a.feb : "",
        a ? a.mar : "",
        a ? a.apr : "",
        a ? a.may : "",
        a ? a.jun : "",
        a ? a.jul : "",
        a ? a.aug : "",
        a ? a.sep : "",
        a ? a.oct : "",
        a ? a.nov : "",
        a ? a.dec : ""
      ]);
    }

    // Blank separator row
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

  // 4. Style Target Rows (Light Blue #D9EEF8 across entire row)
  for (const rIdx of targetRows) {
    const tRange = sheet.getRangeByIndexes(rIdx, 0, 1, mprHeaders.length);
    tRange.getFormat().getFill().setColor("#D9EEF8");
  }

  // 5. Format Percentage values
  for (const rIdx of percentRows) {
    const valRange = sheet.getRangeByIndexes(rIdx, 10, 1, 12);
    valRange.setNumberFormatLocal("0.0%");
  }

  // Auto-fit columns
  sheet.getUsedRange()?.getFormat().autofitColumns();

  return tableData.length;
}
