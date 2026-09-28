/**
 * SCRIPT 1: Extract Functional DMB Review Data (Office Script / TypeScript)
 * =========================================================================
 * Location: Run on 'Functional DMB Review Sheets-17th_sept.xlsx' in SharePoint / Excel Online.
 * Action in Power Automate: "Run script" -> Select Functional DMB workbook.
 * Returns: JSON string containing extracted functional KPI rows with exact metadata & values.
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
}

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
    if (sNameLower.includes("mpr review") || sNameLower.includes("sheet1")) continue;

    // Check if sheet is a Functional tab
    const isFunctional = functionalKeywords.some(k => sNameLower.includes(k));
    if (!isFunctional && !sNameLower.includes("mastersheet")) continue;

    const used = sheet.getUsedRange();
    if (!used) continue;
    const values = used.getValues();
    if (values.length < 5) continue;

    const functionName = cleanFunctionName(sName);

    // Locate header row
    let headerRowIdx = 4;
    for (let r = 0; r < Math.min(10, values.length); r++) {
      const rowStr = values[r].map(v => String(v).toLowerCase()).join(" ");
      if (rowStr.includes("kpi name") || rowStr.includes("mandatory kpi") || rowStr.includes("target/ actual") || rowStr.includes("jan")) {
        headerRowIdx = r;
        break;
      }
    }

    const headerRow = values[headerRowIdx];
    const monthCols = getMonthColumnIndexes(headerRow);

    let curKpi = "";
    let curDef = "";
    let curOp = ">=";
    let curAop: number | string = "";
    let curUnit = "%";
    let curNat = "Higher the better";
    let curFreq = "Monthly";

    for (let r = headerRowIdx + 1; r < values.length; r++) {
      const row = values[r];
      const firstCell = String(row[0] || "").trim();

      // Check if this row defines a new KPI
      if (firstCell && !firstCell.toLowerCase().startsWith("value lever") && !firstCell.startsWith("#REF") && firstCell !== "None") {
        curKpi = firstCell;
        curDef = String(row[1] || "").trim();
        curOp = String(row[2] || ">=").trim();
        curAop = row[3] !== undefined && row[3] !== null ? row[3] : "";
        curUnit = String(row[4] || "%").trim();
        curNat = String(row[5] || "Higher the better").trim();
        curFreq = String(row[6] || "Monthly").trim();
      }

      // Check Target or Actual row (Column H / index 7 in standard sheets)
      let typeVal = String(row[7] || "").trim().toUpperCase();
      if (!typeVal && row.length > 10) {
        // In case Target/Actual is in column 10
        typeVal = String(row[10] || "").trim().toUpperCase();
      }

      if ((typeVal === "TARGET" || typeVal === "ACTUAL" || typeVal === "T" || typeVal === "A") && curKpi) {
        const getMonthVal = (mIdx: number): number | string => {
          const cIdx = monthCols[mIdx];
          if (cIdx !== undefined && cIdx < row.length) {
            const v = row[cIdx];
            return v !== null && v !== undefined && String(v) !== "#REF!" && String(v) !== "None" ? v : "";
          }
          return "";
        };

        let dType = "Whole Num";
        if (curUnit === "%") dType = "Percentage";
        else if (curUnit === "Mn" || curUnit === "K" || curUnit.toLowerCase().includes("decimal")) dType = "Decimal";

        extractedRows.push({
          functionName: functionName,
          sourceSheet: sName,
          kpiName: curKpi,
          definition: curDef,
          operator: curOp,
          target2026: curAop,
          unit: curUnit,
          nature: curNat,
          frequency: curFreq,
          rowType: typeVal.startsWith("T") ? "Target" : "Actual",
          jan: getMonthVal(0),
          feb: getMonthVal(1),
          mar: getMonthVal(2),
          apr: getMonthVal(3),
          may: getMonthVal(4),
          jun: getMonthVal(5),
          jul: getMonthVal(6),
          aug: getMonthVal(7),
          sep: getMonthVal(8),
          oct: getMonthVal(9),
          nov: getMonthVal(10),
          dec: getMonthVal(11),
          dataType: dType
        });
      }
    }
  }

  console.log(`Extracted ${extractedRows.length} functional KPI rows.`);
  return JSON.stringify(extractedRows);
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

  // Fallback: If dates were numeric serials / Date objects, calendar columns in Functional review are columns 8 to 19
  for (let m = 0; m < 12; m++) {
    if (mapping[m] === undefined) {
      mapping[m] = 8 + m;
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
