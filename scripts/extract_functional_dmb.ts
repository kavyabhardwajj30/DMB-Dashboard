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

  // 1. Prefer extracting from existing 'MasterSheet' if available
  const masterWs = workbook.getWorksheet("MasterSheet") || workbook.getWorksheet("Mastersheet");
  if (masterWs && masterWs.getUsedRange()) {
    const values = masterWs.getUsedRange().getValues();
    if (values.length > 5) {
      let currentFunction = "Quality";
      let monthCols: { [m: number]: number } = {};
      let isHeaderActive = false;

      for (let r = 0; r < values.length; r++) {
        const row = values[r];
        const firstCell = String(row[0] || "").trim();

        if (firstCell.toLowerCase() === "mastersheet" || !firstCell) continue;

        // Check if row is a section banner (e.g. 'Quality DMB', 'Regulatory')
        const nextFirstCell = r + 1 < values.length ? String(values[r + 1][0] || "").trim().toLowerCase() : "";
        if (nextFirstCell.includes("kpi") || firstCell.toLowerCase().endsWith("dmb")) {
          currentFunction = cleanFunctionName(firstCell);
          continue;
        }

        // Check if row is a column header
        if (firstCell.toLowerCase().startsWith("kpi name") || firstCell.toLowerCase().startsWith("mandatory kpi")) {
          monthCols = getMonthColumnIndexes(row);
          isHeaderActive = true;
          continue;
        }

        // Check Target / Actual row
        const typeCell = String(row[7] || "").trim();
        const typeUpper = typeCell.toUpperCase();
        if (typeUpper === "TARGET" || typeUpper === "ACTUAL" || typeUpper === "T" || typeUpper === "A") {
          const kpiName = firstCell;
          if (!kpiName || kpiName.startsWith("#REF")) continue;

          const getMonthVal = (mIdx: number): number | string => {
            const cIdx = monthCols[mIdx];
            if (cIdx !== undefined && cIdx < row.length) {
              const v = row[cIdx];
              return v !== null && v !== undefined && String(v) !== "#REF!" ? v : "";
            }
            return "";
          };

          const unitVal = String(row[4] || "").trim();
          let dataType = "Whole Num";
          if (unitVal === "%") dataType = "Percentage";
          else if (unitVal === "Mn" || unitVal === "K") dataType = "Decimal";

          extractedRows.push({
            functionName: currentFunction,
            sourceSheet: "MasterSheet",
            kpiName: kpiName,
            definition: String(row[1] || "").trim(),
            operator: String(row[2] || ">=").trim(),
            target2026: row[3] !== undefined && row[3] !== null ? row[3] : "",
            unit: unitVal,
            nature: String(row[5] || "").trim(),
            frequency: String(row[6] || "Monthly").trim(),
            rowType: typeUpper.startsWith("T") ? "Target" : "Actual",
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
            dataType: dataType
          });
        }
      }

      if (extractedRows.length > 0) {
        console.log(`Extracted ${extractedRows.length} rows directly from MasterSheet.`);
        return JSON.stringify(extractedRows);
      }
    }
  }

  // 2. Fallback: Extract from individual functional sheets
  const functionalKeywords = [
    "quality", "regulatory", "isc", "procurement",
    "r&d", "marketing", "customer service", "commercial excellence",
    "nar", "europe", "growth", "finance"
  ];

  for (const sheet of workbook.getWorksheets()) {
    const sName = sheet.getName().trim();
    const sNameLower = sName.toLowerCase();
    if (sNameLower.includes("mastersheet") || sNameLower.includes("mpr")) continue;

    const isMatch = functionalKeywords.some(k => sNameLower.includes(k));
    if (!isMatch) continue;

    const functionName = cleanFunctionName(sName);
    const used = sheet.getUsedRange();
    if (!used) continue;
    const values = used.getValues();
    if (values.length < 5) continue;

    let headerRowIdx = 4;
    for (let r = 0; r < Math.min(10, values.length); r++) {
      const rowStr = values[r].map(v => String(v).toLowerCase()).join(" ");
      if (rowStr.includes("jan") || rowStr.includes("target/ actual") || rowStr.includes("target aop")) {
        headerRowIdx = r;
        break;
      }
    }

    const headerRow = values[headerRowIdx];
    const monthCols = getMonthColumnIndexes(headerRow);

    for (let r = headerRowIdx + 1; r < values.length; r++) {
      const row = values[r];
      const typeVal = String(row[7] || "").trim().toUpperCase();
      if (typeVal === "TARGET" || typeVal === "ACTUAL" || typeVal === "T" || typeVal === "A") {
        const kpiName = String(row[0] || "").trim();
        if (!kpiName || kpiName.startsWith("Value Lever") || kpiName.startsWith("#REF")) continue;

        const getMonthVal = (mIdx: number): number | string => {
          const cIdx = monthCols[mIdx];
          if (cIdx !== undefined && cIdx < row.length) {
            const v = row[cIdx];
            return v !== null && v !== undefined && String(v) !== "#REF!" ? v : "";
          }
          return "";
        };

        const unitVal = String(row[4] || "").trim();
        extractedRows.push({
          functionName: functionName,
          sourceSheet: sName,
          kpiName: kpiName,
          definition: String(row[1] || "").trim(),
          operator: String(row[2] || ">=").trim(),
          target2026: row[3] !== undefined && row[3] !== null ? row[3] : "",
          unit: unitVal,
          nature: String(row[5] || "").trim(),
          frequency: String(row[6] || "Monthly").trim(),
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
          dataType: unitVal === "%" ? "Percentage" : "Whole Num"
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
  // Default fallback to columns 8..19 if not found
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
