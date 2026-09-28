/**
 * SCRIPT 2: Extract Strategic Execution AOP Critical Data (Office Script / TypeScript)
 * ====================================================================================
 * Location: Run on 'Strategic Execution Dashboard-17Th_sept.xlsx' in SharePoint / Excel Online.
 * Action in Power Automate: "Run script" -> Select Strategic Execution workbook.
 * Returns: JSON string containing extracted strategic KPI rows.
 */

interface StrategicKPIRow {
  strategicImperative: string;
  sourceSheet: string;
  kpiName: string;
  owner: string;
  definition: string;
  target2026: number | string;
  column1?: string;
  column2?: string;
  nature: string;
  unit: string;
  frequency: string;
  rowType: "T" | "A";
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
  console.log("Starting strategic execution data extraction...");
  const extractedRows: StrategicKPIRow[] = [];

  let aopSheet = workbook.getWorksheet("AOP Critical") || workbook.getWorksheet("Mastersheet");
  if (!aopSheet) {
    const sheets = workbook.getWorksheets();
    aopSheet = sheets.find(s => s.getName().toLowerCase().includes("aop") || s.getName().toLowerCase().includes("strategic")) || sheets[0];
  }

  if (!aopSheet) return JSON.stringify([]);

  const usedRange = aopSheet.getUsedRange();
  if (!usedRange) return JSON.stringify([]);
  const values = usedRange.getValues();
  if (values.length < 5) return JSON.stringify([]);

  // Detect header row
  let headerRowIndex = 4;
  for (let r = 0; r < Math.min(10, values.length); r++) {
    const rowStr = values[r].map(v => String(v).toLowerCase()).join(" ");
    if (rowStr.includes("core kpi") || rowStr.includes("metric owner") || rowStr.includes("jan")) {
      headerRowIndex = r;
      break;
    }
  }

  const headerRow = values[headerRowIndex];
  const monthCols = getStrategicMonthCols(headerRow);

  let curImperative = "Patient Safety & Quality";
  let curKpi = "";
  let curOwner = "";
  let curDef = "";
  let curAop: number | string = "";
  let curCol1 = "";
  let curCol2 = "";
  let curNat = "Higher the better";
  let curUnit = "%";
  let curFreq = "Monthly";

  for (let r = headerRowIndex + 1; r < values.length; r++) {
    const row = values[r];
    const impVal = String(row[0] || "").trim();
    if (impVal && !impVal.startsWith("MoS") && !impVal.startsWith("Strategic") && !impVal.startsWith("#REF") && impVal !== "None") {
      curImperative = normalizeImperativeName(impVal);
    }

    const kpiVal = String(row[1] || "").trim();
    if (kpiVal && !kpiVal.toLowerCase().startsWith("value lever") && !kpiVal.startsWith("#REF") && kpiVal !== "None") {
      curKpi = kpiVal;
      curOwner = String(row[2] || "").trim();
      curDef = String(row[3] || "").trim();
      curAop = row[4] !== undefined && row[4] !== null ? row[4] : "";
      curCol1 = String(row[5] || "").trim();
      curCol2 = String(row[6] || "").trim();
      curNat = String(row[7] || "Higher the better").trim();
      curUnit = String(row[8] || "%").trim();
      curFreq = String(row[9] || "Monthly").trim();
    }

    // Check Type (Column K / index 10 or Column J / index 9)
    let typeVal = String(row[10] || "").trim().toUpperCase();
    if (!typeVal && row.length > 9) {
      typeVal = String(row[9] || "").trim().toUpperCase();
    }

    if ((typeVal === "T" || typeVal === "A" || typeVal === "TARGET" || typeVal === "ACTUAL") && curKpi) {
      const getMonthVal = (mIdx: number): number | string => {
        const cIdx = monthCols[mIdx];
        if (cIdx !== undefined && cIdx < row.length) {
          const v = row[cIdx];
          return v !== null && v !== undefined && String(v) !== "#REF!" && String(v) !== "None" ? v : "";
        }
        return "";
      };

      extractedRows.push({
        strategicImperative: curImperative,
        sourceSheet: aopSheet.getName(),
        kpiName: curKpi,
        owner: curOwner,
        definition: curDef,
        target2026: curAop,
        column1: curCol1,
        column2: curCol2,
        nature: curNat,
        unit: curUnit,
        frequency: curFreq,
        rowType: typeVal.startsWith("T") ? "T" : "A",
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
        dataType: curUnit === "%" ? "Percentage" : "Whole Num"
      });
    }
  }

  console.log(`Extracted ${extractedRows.length} strategic KPI rows.`);
  return JSON.stringify(extractedRows);
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

  // Fallback: If dates were serials, months in AOP Critical are columns 11 to 22 (or 10 to 21 in Mastersheet)
  for (let m = 0; m < 12; m++) {
    if (mapping[m] === undefined) {
      mapping[m] = 11 + m;
    }
  }

  return mapping;
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
