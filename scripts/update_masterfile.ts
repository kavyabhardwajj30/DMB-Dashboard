/**
 * SCRIPT 3: Populate Masterfile Worksheets (Office Script / TypeScript)
 * =====================================================================
 * Location: Run on 'Masterfile_DMB_Dashboard.xlsx' in SharePoint / Excel Online.
 * Action in Power Automate: "Run script" -> Select Masterfile workbook.
 * Parameters:
 *   - functionalJson (string): JSON output from SCRIPT 1 (extract_functional_dmb)
 *   - strategicJson (string): JSON output from SCRIPT 2 (extract_strategic_execution)
 *
 * Flow: when either source workbook changes, Power Automate runs SCRIPT 1 and
 * SCRIPT 2 to read them, then runs this script to rewrite the two Masterfile
 * tabs with the fresh data in the Mastersheet layout and styling:
 *   Functional DMB Review Sheets  -> 'DMB Masterfile'
 *   Strategic Execution Dashboard -> 'MPR Masterfile'
 *
 * Naming rule: when the same KPI name appears in more than one function tab,
 * the function name is appended, e.g. "CTB (12 weeks)" in both the ISC and the
 * Procurement tab is written as "CTB (12 weeks)(ISC)" and
 * "CTB (12 weeks)(Procurement)". The same applies when one tab repeats a KPI
 * name: the KPI definition is then used to tell the two apart.
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
  category?: string;
}

/** One KPI: its metadata plus its Target row and its Actual row. */
interface KPIEntry {
  section: string;
  groupName: string; // function name (DMB) or strategic imperative (MPR)
  kpiName: string;
  displayName: string;
  owner: string;
  definition: string;
  operator: string;
  target2026: number | string;
  column1: string;
  column2: string;
  nature: string;
  unit: string;
  frequency: string;
  dataType: string;
  category: string;
  target?: KPIRow;
  actual?: KPIRow;
}

// =====================================================================
// CONFIGURATION
// =====================================================================

/** Function tabs that share one Masterfile section (ISC + Procurement). */
const FUNCTION_TO_SECTION: { [functionName: string]: string } = {
  "quality": "Quality DMB",
  "regulatory": "Regulatory",
  "isc": "ISC & Procurement",
  "procurement": "ISC & Procurement",
  "r&d": "R&D",
  "marketing": "Marketing",
  "customer service": "Customer Service",
  "commercial excellence": "Commercial Excellence",
  "nar": "NAR",
  "europe": "Europe",
  "growth": "Growth",
  "finance": "Finance"
};

/** Section order in 'DMB Masterfile'. Anything else is appended at the end. */
const DMB_SECTION_ORDER: string[] = [
  "Quality DMB",
  "Regulatory",
  "ISC & Procurement",
  "R&D",
  "Marketing",
  "Customer Service",
  "Commercial Excellence",
  "NAR",
  "Europe",
  "Growth",
  "Finance"
];

/** Section order in 'MPR Masterfile'. */
const MPR_SECTION_ORDER: string[] = [
  "Patient Safety & Quality",
  "Customer Focus",
  "Improve Deliverability & Profitability",
  "Roadmap competitiveness & Innovation agility",
  "Drive growth through Commercial Excellence",
  "Environmental, Social & Governance"
];

/** Function tabs to leave out of the Masterfile, e.g. ["Finance"]. */
const SKIP_FUNCTIONS: string[] = [];

/**
 * Preferred KPI name for a given KPI definition. Used when one tab repeats a
 * KPI name, so the Masterfile keeps the wording the business already uses.
 * Add an entry as "definition text": "KPI name to write".
 */
const NAME_OVERRIDES: { [definition: string]: string } = {
  "scar clsoure % scars closed < 180d": "SCAR Performance (closed less then 180d)",
  "scars on time to plan (ottp)": "SCAR Performance (OTTP)",
  "oit growth % (eq)": "OIT Growth % (EQ)",
  "oit growth % (cs)": "OIT Growth % (CS)"
};

/*
 * Mastersheet styling, taken from Masterfile_DMB_Dashboard.xlsx:
 *   - Aptos Narrow 11 everywhere, black text.
 *   - Row 1 title in white bold on the blue accent, 14pt on the DMB tab and
 *     12pt on the MPR tab.
 *   - Section names (e.g. "Quality DMB") in bold.
 *   - Header rows plain for the KPI columns, and white bold on blue with thin
 *     top and bottom borders from the first month column onwards.
 *   - Target rows carry a light blue band with thin borders across the months,
 *     Actual rows the same borders without the fill.
 *   - Month cells use a number format that fits the values of that row.
 */
const BODY_FONT = "Aptos Narrow";
const BODY_FONT_SIZE = 11;
const TEXT_COLOR = "#000000";
const TITLE_FILL = "#4F81BD";        // theme accent 1
const TITLE_FONT_COLOR = "#FFFFFF";
const HEADER_FILL = "#4F81BD";
const HEADER_FONT_COLOR = "#FFFFFF";
const TARGET_BAND_FILL = "#C0E6F5";
const BORDER_COLOR = "#000000";
const DMB_TITLE_FONT_SIZE = 14;
const MPR_TITLE_FONT_SIZE = 12;
const DMB_TITLE_ROW_HEIGHT = 18.6;
const MPR_TITLE_ROW_HEIGHT = 16;
const DMB_FROZEN_COLUMNS = 4;        // the DMB tab freezes columns A to D
const GENERAL_FORMAT = "General";
const MAX_DECIMALS = 2;

const DMB_HEADERS: string[] = [
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
  "Data Type",
  "KPI category"
];

const MPR_HEADERS: string[] = [
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
  "Mar",
  "Apr",
  "May",
  "Jun",
  "Jul",
  "Aug",
  "Sep",
  "Oct",
  "Nov",
  "Dec",
  "Data Type"
];

/*
 * Column widths in points, matching the Mastersheet.
 * (Excel character width -> points is (characters * 7 + 5) * 0.75.)
 * DMB: KPI Name, Definition, Operators, Target AOP, Units, Metric nature,
 *      Frequency, Target/Actual, then Jan..Dec, Data Type, KPI category.
 */
const DMB_COLUMN_WIDTHS: { [columnIndex: number]: number } = {
  0: 232, 1: 191, 2: 62, 3: 64, 4: 53, 5: 85, 6: 66, 7: 70,
  8: 64, 9: 64, 10: 64, 11: 64, 12: 64, 13: 64,
  14: 64, 15: 64, 16: 64, 17: 64, 18: 64, 19: 64,
  20: 70, 21: 98
};
/*
 * MPR: Core KPI's, Metric Owner, Definition, AOP 2026, Column1, Column2,
 *      Metric nature, Units, Frequency, Target/Actual, Jan..Dec, Data Type.
 */
const MPR_COLUMN_WIDTHS: { [columnIndex: number]: number } = {
  0: 214, 1: 98, 2: 294, 3: 49, 4: 60, 5: 60, 6: 86, 7: 49, 8: 70, 9: 83,
  10: 49, 11: 49, 12: 49, 13: 49, 14: 49, 15: 49,
  16: 49, 17: 49, 18: 49, 19: 49, 20: 49, 21: 49,
  22: 79
};

// =====================================================================
// MAIN
// =====================================================================

function main(
  workbook: ExcelScript.Workbook,
  functionalJson?: string,
  strategicJson?: string
): { dmbRows: number; mprRows: number; status: string } {
  console.log("Starting Masterfile sync with exact Mastersheet styling...");

  let dmbRowsCount = 0;
  let mprRowsCount = 0;

  // 1. Process Functional Review Data into 'DMB Masterfile'
  if (functionalJson && functionalJson.trim().length > 2) {
    try {
      const functionalRows: KPIRow[] = JSON.parse(functionalJson);
      if (functionalRows.length > 0) {
        let dmbSheet = workbook.getWorksheet("DMB Masterfile");
        if (!dmbSheet) {
          dmbSheet = workbook.addWorksheet("DMB Masterfile");
        }
        dmbRowsCount = writeDMBMasterfileExact(dmbSheet, functionalRows);
        console.log(`Successfully formatted and wrote ${dmbRowsCount} rows to DMB Masterfile.`);
      } else {
        console.warn("functionalJson parsed to empty array [].");
      }
    } catch (err) {
      console.error("Error writing DMB Masterfile:", err);
    }
  }

  // 2. Process Strategic Execution Data into 'MPR Masterfile'
  if (strategicJson && strategicJson.trim().length > 2) {
    try {
      const strategicRows: KPIRow[] = JSON.parse(strategicJson);
      if (strategicRows.length > 0) {
        let mprSheet = workbook.getWorksheet("MPR Masterfile");
        if (!mprSheet) {
          mprSheet = workbook.addWorksheet("MPR Masterfile");
        }
        mprRowsCount = writeMPRMasterfileExact(mprSheet, strategicRows);
        console.log(`Successfully formatted and wrote ${mprRowsCount} rows to MPR Masterfile.`);
      } else {
        console.warn("strategicJson parsed to empty array [].");
      }
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

// =====================================================================
// SHARED HELPERS
// =====================================================================

function cleanText(value: string | number | boolean | undefined | null): string {
  if (value === undefined || value === null) return "";
  return String(value).replace(/ /g, " ").trim();
}

/** Lookup key for NAME_OVERRIDES: lower case, line breaks folded to spaces. */
function definitionKey(definition: string): string {
  return cleanText(definition).replace(/\s+/g, " ").toLowerCase();
}

function isBlank(value: number | string | undefined | null): boolean {
  const text = cleanText(value as string);
  return text === "" || text === "None" || text.indexOf("#REF") === 0;
}

function monthValues(row: KPIRow | undefined): (string | number)[] {
  if (!row) return ["", "", "", "", "", "", "", "", "", "", "", ""];
  const raw = [row.jan, row.feb, row.mar, row.apr, row.may, row.jun,
               row.jul, row.aug, row.sep, row.oct, row.nov, row.dec];
  return raw.map(v => (isBlank(v) ? "" : v));
}

/** True when neither the Target nor the Actual row holds a single value. */
function hasNoValues(entry: KPIEntry): boolean {
  const all = monthValues(entry.target).concat(monthValues(entry.actual));
  return all.every(v => v === "");
}

function rowIsTarget(row: KPIRow): boolean {
  return String(row.rowType).toUpperCase().indexOf("T") === 0;
}

/** Text inside the last (...) of a definition, e.g. "OIT Growth % (EQ)" -> "EQ". */
function parentheticalTag(definition: string): string {
  const matches = cleanText(definition).match(/\(([^()]+)\)\s*$/);
  return matches ? matches[1].trim() : "";
}

/** Words of this definition that the sibling definitions do not have. */
function distinctiveTag(definition: string, siblingDefinitions: string[]): string {
  const shared: { [word: string]: boolean } = {};
  for (const sibling of siblingDefinitions) {
    for (const word of cleanText(sibling).toLowerCase().split(/\s+/)) {
      shared[word] = true;
    }
  }
  const words = cleanText(definition).split(/\s+/)
    .filter(word => word.length > 0 && !shared[word.toLowerCase()]);
  return words.slice(0, 3).join(" ");
}

/**
 * Decide the KPI name written into the Masterfile.
 * - A KPI name used by more than one function tab gets "(Function)" appended.
 * - A KPI name repeated inside one section is told apart by its definition.
 */
function resolveDisplayNames(entries: KPIEntry[]): void {
  const groupsByName: { [kpiName: string]: string[] } = {};
  for (const entry of entries) {
    const key = entry.kpiName.toLowerCase();
    if (!groupsByName[key]) groupsByName[key] = [];
    if (groupsByName[key].indexOf(entry.groupName) < 0) {
      groupsByName[key].push(entry.groupName);
    }
  }

  for (const entry of entries) {
    const override = NAME_OVERRIDES[definitionKey(entry.definition)];
    if (override) {
      entry.displayName = override;
      continue;
    }
    // Same KPI name in two or more function tabs -> add the function name.
    entry.displayName = groupsByName[entry.kpiName.toLowerCase()].length > 1
      ? `${entry.kpiName}(${entry.groupName})`
      : entry.kpiName;
  }

  // Still duplicated inside a section? Tell them apart by their definition.
  const seen: { [key: string]: KPIEntry[] } = {};
  for (const entry of entries) {
    const key = `${entry.section}||${entry.displayName.toLowerCase()}`;
    if (!seen[key]) seen[key] = [];
    seen[key].push(entry);
  }

  for (const key of Object.keys(seen)) {
    const clashing = seen[key];
    if (clashing.length < 2) continue;
    for (let i = 0; i < clashing.length; i++) {
      const entry = clashing[i];
      const others = clashing.filter(other => other !== entry).map(other => other.definition);
      const tag = parentheticalTag(entry.definition)
        || distinctiveTag(entry.definition, others)
        || String(i + 1);
      entry.displayName = `${entry.displayName} (${tag})`;
    }
  }
}

/** Pair the Target and Actual rows of each KPI, keeping the source order. */
function buildEntries(rows: KPIRow[], isStrategic: boolean): KPIEntry[] {
  const entries: KPIEntry[] = [];
  const index: { [key: string]: KPIEntry } = {};

  for (const row of rows) {
    const kpiName = cleanText(row.kpiName);
    if (!kpiName || kpiName.indexOf("#REF") === 0) continue;

    const groupName = isStrategic
      ? cleanText(row.strategicImperative) || "Strategic Imperatives"
      : cleanText(row.functionName) || "General";

    if (!isStrategic && SKIP_FUNCTIONS.indexOf(groupName) >= 0) continue;

    const section = isStrategic
      ? groupName
      : (FUNCTION_TO_SECTION[groupName.toLowerCase()] || groupName);

    const definition = cleanText(row.definition);
    /*
     * The function tab is part of the identity: ISC and Procurement both hold
     * a "CTB (12 weeks)" with the same definition, and they are two different
     * KPIs even though they share one Masterfile section.
     */
    const key = `${section}||${groupName.toLowerCase()}||${kpiName.toLowerCase()}||${definition.toLowerCase()}`;

    if (!index[key]) {
      index[key] = {
        section: section,
        groupName: groupName,
        kpiName: kpiName,
        displayName: kpiName,
        owner: cleanText(row.owner),
        definition: definition,
        operator: cleanText(row.operator),
        target2026: isBlank(row.target2026) ? "" : row.target2026,
        column1: cleanText(row.column1),
        column2: cleanText(row.column2),
        nature: cleanText(row.nature),
        unit: cleanText(row.unit),
        frequency: cleanText(row.frequency),
        dataType: cleanText(row.dataType) || dataTypeForUnit(cleanText(row.unit)),
        category: cleanText(row.category)
      };
      entries.push(index[key]);
    }

    const entry = index[key];
    if (rowIsTarget(row)) {
      entry.target = row;
    } else {
      entry.actual = row;
    }
    // Metadata sits on the Target row in the source sheets; fill any gaps.
    if (!entry.owner) entry.owner = cleanText(row.owner);
    if (!entry.frequency) entry.frequency = cleanText(row.frequency);
    if (!entry.category) entry.category = cleanText(row.category);
  }

  // Drop template rows that carry no Target and no Actual value at all.
  const populated = entries.filter(entry => !(hasNoValues(entry) && !entry.definition));
  resolveDisplayNames(populated);
  return populated;
}

function dataTypeForUnit(unit: string): string {
  const lower = unit.toLowerCase();
  if (lower === "%") return "Percentage";
  if (lower === "mn" || lower === "k" || lower.indexOf("decimal") >= 0) return "Decimal";
  return "Whole Number";
}

function sortSections(sections: string[], order: string[]): string[] {
  const known = order.filter(section => sections.indexOf(section) >= 0);
  const extra = sections.filter(section => order.indexOf(section) < 0).sort();
  return known.concat(extra);
}

function sectionsOf(entries: KPIEntry[]): string[] {
  const sections: string[] = [];
  for (const entry of entries) {
    if (sections.indexOf(entry.section) < 0) sections.push(entry.section);
  }
  return sections;
}

/** The generated block, with the row kinds that styling is applied to. */
interface SheetLayout {
  data: (string | number)[][];
  sectionRows: number[];
  headerRows: number[];
  targetRows: number[];
  /** Header row through last KPI row of a section, for the month grid lines. */
  borderBlocks: { start: number; end: number }[];
  rowFormats: { [rowIndex: number]: string };
}

/** Where the months sit and how the sheet is sized. */
interface SheetStyleOptions {
  title: string;
  titleFontSize: number;
  titleRowHeight: number;
  monthStartColumn: number;
  columnCount: number;
  columnWidths: { [columnIndex: number]: number };
  frozenColumns: number;
}

/**
 * Pick the number format of one row from its own values, the way the
 * Mastersheet does: whole percentages as 0%, 87.5% as 0.0%, 54.51% as 0.00%,
 * counts as 0 and amounts as 0.00.
 */
function numberFormatFor(values: (string | number)[], unit: string): string {
  const isPercent = cleanText(unit) === "%";
  let decimals = 0;
  let sawNumber = false;

  for (const value of values) {
    if (typeof value !== "number" || !isFinite(value)) continue;
    sawNumber = true;
    const shown = isPercent ? value * 100 : value;
    let needed = MAX_DECIMALS;
    for (let places = 0; places <= MAX_DECIMALS; places++) {
      const factor = Math.pow(10, places);
      if (Math.abs(shown * factor - Math.round(shown * factor)) < 0.000001) {
        needed = places;
        break;
      }
    }
    if (needed > decimals) decimals = needed;
  }

  if (!sawNumber) return GENERAL_FORMAT;
  const tail = decimals === 0 ? "" : (decimals === 1 ? ".0" : ".00");
  return isPercent ? `0${tail}%` : `0${tail}`;
}

/** Thin black borders on the given edges of a range. */
function setBorders(range: ExcelScript.Range, edges: ExcelScript.BorderIndex[]): void {
  for (const edge of edges) {
    const border = range.getFormat().getRangeBorder(edge);
    border.setStyle(ExcelScript.BorderLineStyle.continuous);
    border.setWeight(ExcelScript.BorderWeight.thin);
    border.setColor(BORDER_COLOR);
  }
}

/** Apply the Mastersheet look to the block that was just written. */
function applyMastersheetStyling(
  sheet: ExcelScript.Worksheet,
  layout: SheetLayout,
  options: SheetStyleOptions
): void {
  const rowCount = layout.data.length;
  const columnCount = options.columnCount;
  const monthStart = options.monthStartColumn;
  const bandWidth = columnCount - monthStart; // months plus the trailing columns

  // 1. Aptos Narrow 11, black, across the whole block.
  const whole = sheet.getRangeByIndexes(0, 0, rowCount, columnCount);
  whole.getFormat().getFont().setName(BODY_FONT);
  whole.getFormat().getFont().setSize(BODY_FONT_SIZE);
  whole.getFormat().getFont().setColor(TEXT_COLOR);
  whole.getFormat().getFont().setBold(false);

  // 2. Row 1 title on the blue accent.
  const titleCell = sheet.getRangeByIndexes(0, 0, 1, 1);
  titleCell.getFormat().getFont().setSize(options.titleFontSize);
  titleCell.getFormat().getFont().setBold(true);
  titleCell.getFormat().getFont().setColor(TITLE_FONT_COLOR);
  titleCell.getFormat().getFill().setColor(TITLE_FILL);
  titleCell.getFormat().setRowHeight(options.titleRowHeight);

  // 3. Section names in bold.
  for (const rowIndex of layout.sectionRows) {
    sheet.getRangeByIndexes(rowIndex, 0, 1, 1).getFormat().getFont().setBold(true);
  }

  // 4. Header rows: white bold on blue from the first month column onwards.
  for (const rowIndex of layout.headerRows) {
    const band = sheet.getRangeByIndexes(rowIndex, monthStart, 1, bandWidth);
    band.getFormat().getFont().setBold(true);
    band.getFormat().getFont().setColor(HEADER_FONT_COLOR);
    band.getFormat().getFill().setColor(HEADER_FILL);
    setBorders(band, [
      ExcelScript.BorderIndex.edgeTop,
      ExcelScript.BorderIndex.edgeBottom,
      ExcelScript.BorderIndex.edgeRight
    ]);
  }

  // 5. A line under every month row, from the header down to the last KPI.
  for (const block of layout.borderBlocks) {
    const rows = block.end - block.start + 1;
    setBorders(sheet.getRangeByIndexes(block.start, monthStart, rows, 12), [
      ExcelScript.BorderIndex.edgeTop,
      ExcelScript.BorderIndex.edgeBottom,
      ExcelScript.BorderIndex.insideHorizontal
    ]);
  }

  // 6. Target rows carry the light blue band across the months.
  for (const rowIndex of layout.targetRows) {
    sheet.getRangeByIndexes(rowIndex, monthStart, 1, 12)
      .getFormat().getFill().setColor(TARGET_BAND_FILL);
  }

  // 7. Number format per row, chosen from that row's values.
  for (const key of Object.keys(layout.rowFormats)) {
    const rowIndex = Number(key);
    sheet.getRangeByIndexes(rowIndex, monthStart, 1, 12)
      .setNumberFormatLocal(layout.rowFormats[rowIndex]);
  }

  // 8. Column widths and frozen columns.
  for (const key of Object.keys(options.columnWidths)) {
    const columnIndex = Number(key);
    if (columnIndex >= columnCount) continue;
    sheet.getRangeByIndexes(0, columnIndex, 1, 1)
      .getFormat().setColumnWidth(options.columnWidths[columnIndex]);
  }
  if (options.frozenColumns > 0) {
    sheet.getFreezePanes().freezeColumns(options.frozenColumns);
  }

  console.log(`Applied Mastersheet styling to '${options.title}'.`);
}

/**
 * Build the rows of one tab: title banner, then a section name, a header row
 * and a Target / Actual pair per KPI.
 */
function buildLayout(
  entries: KPIEntry[],
  headers: string[],
  isStrategic: boolean
): SheetLayout {
  const columnCount = headers.length;
  const layout: SheetLayout = {
    data: [],
    sectionRows: [],
    headerRows: [],
    targetRows: [],
    borderBlocks: [],
    rowFormats: {}
  };

  const blankRow = (): (string | number)[] => {
    const row: (string | number)[] = [];
    for (let i = 0; i < columnCount; i++) row.push("");
    return row;
  };

  const banner = blankRow();
  banner[0] = isStrategic ? "Strategic Imperatives" : "Mastersheet";
  layout.data.push(banner);
  layout.data.push(blankRow());

  const order = isStrategic ? MPR_SECTION_ORDER : DMB_SECTION_ORDER;
  for (const section of sortSections(sectionsOf(entries), order)) {
    const sectionRow = blankRow();
    sectionRow[0] = section;
    layout.sectionRows.push(layout.data.length);
    layout.data.push(sectionRow);

    const headerRowIndex = layout.data.length;
    layout.headerRows.push(headerRowIndex);
    layout.data.push(headers.slice() as (string | number)[]);

    for (const entry of entries.filter(item => item.section === section)) {
      const rowTypes = isStrategic ? ["T", "A"] : ["Target", "Actual"];
      for (const rowType of rowTypes) {
        const isTarget = rowType === "T" || rowType === "Target";
        const monthly = monthValues(isTarget ? entry.target : entry.actual);
        const rowIndex = layout.data.length;

        if (isTarget) {
          layout.targetRows.push(rowIndex);
        }
        layout.rowFormats[rowIndex] = numberFormatFor(monthly, entry.unit);

        const head: (string | number)[] = isStrategic
          ? [
              entry.displayName,
              entry.owner,
              entry.definition,
              entry.target2026,
              entry.column1,
              entry.column2,
              entry.nature,
              entry.unit,
              entry.frequency,
              rowType
            ]
          : [
              entry.displayName,
              entry.definition,
              entry.operator,
              entry.target2026,
              entry.unit,
              entry.nature,
              entry.frequency,
              rowType
            ];
        const tail: (string | number)[] = isStrategic
          ? [entry.dataType]
          : [entry.dataType, entry.category];

        layout.data.push(head.concat(monthly).concat(tail));
      }
    }

    if (layout.data.length > headerRowIndex + 1) {
      layout.borderBlocks.push({
        start: headerRowIndex,
        end: layout.data.length - 1
      });
    }
  }

  return layout;
}

// =====================================================================
// DMB MASTERFILE ('Functional DMB Review Sheets' -> function tabs)
// =====================================================================

function writeDMBMasterfileExact(sheet: ExcelScript.Worksheet, rows: KPIRow[]): number {
  const entries = buildEntries(rows, false);
  if (entries.length === 0) return 0;

  sheet.getUsedRange()?.clear(ExcelScript.ClearApplyTo.all);

  const layout = buildLayout(entries, DMB_HEADERS, false);
  const columnCount = DMB_HEADERS.length;
  sheet.getRangeByIndexes(0, 0, layout.data.length, columnCount).setValues(layout.data);

  applyMastersheetStyling(sheet, layout, {
    title: "DMB Masterfile",
    titleFontSize: DMB_TITLE_FONT_SIZE,
    titleRowHeight: DMB_TITLE_ROW_HEIGHT,
    monthStartColumn: 8, // column I
    columnCount: columnCount,
    columnWidths: DMB_COLUMN_WIDTHS,
    frozenColumns: DMB_FROZEN_COLUMNS
  });

  return layout.data.length;
}

// =====================================================================
// MPR MASTERFILE ('Strategic Execution Dashboard' -> AOP Critical)
// =====================================================================

function writeMPRMasterfileExact(sheet: ExcelScript.Worksheet, rows: KPIRow[]): number {
  const entries = buildEntries(rows, true);
  if (entries.length === 0) return 0;

  sheet.getUsedRange()?.clear(ExcelScript.ClearApplyTo.all);

  const layout = buildLayout(entries, MPR_HEADERS, true);
  const columnCount = MPR_HEADERS.length;
  sheet.getRangeByIndexes(0, 0, layout.data.length, columnCount).setValues(layout.data);

  applyMastersheetStyling(sheet, layout, {
    title: "MPR Masterfile",
    titleFontSize: MPR_TITLE_FONT_SIZE,
    titleRowHeight: MPR_TITLE_ROW_HEIGHT,
    monthStartColumn: 10, // column K
    columnCount: columnCount,
    columnWidths: MPR_COLUMN_WIDTHS,
    frozenColumns: 0
  });

  return layout.data.length;
}
