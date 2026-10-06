/**
 * SYNC: Functional DMB + Strategic Execution -> Masterfile (Office Script / TypeScript)
 * ====================================================================================
 * Location: Run on 'Masterfile_DMB_Dashboard.xlsx' in SharePoint / Excel Online.
 * One script syncs both Masterfile tabs:
 *   Functional DMB Review Sheets   -> 'DMB Masterfile'
 *   Strategic Execution Dashboard  -> 'MPR Masterfile'
 *
 * Power Automate (one flow):
 *   1. Trigger   : SharePoint "When a file is created or modified (properties only)"
 *                  on the library/folder that holds both source workbooks.
 *   2. Run script: extract_source.ts on the Functional DMB workbook
 *   3. Run script: extract_source.ts on the Strategic Execution workbook
 *   4. Run script: THIS script on the Masterfile with
 *        functionalJson = result of step 2
 *        strategicJson  = result of step 3
 *   Every change found is written straight to the Masterfile.
 *   An Office Script can only read the workbook it runs on, so steps 2 and 3
 *   are what read the two source files. Rows are routed by their content, so
 *   swapping the two results does no harm; a tab with no rows is skipped.
 *
 * Only what differs is changed:
 *   - a changed cell (value, target, owner, definition, unit, ...) -> only that cell is written
 *   - a KPI added in the source -> a Target/Actual row pair is inserted in the right
 *     section, at the same position as in the source, inside the section's table
 *     and with the formatting of the neighbouring KPI rows
 *   - a KPI removed from the source -> its two rows are deleted
 *   - a new function tab / strategic imperative -> a new section (title, header, table)
 * Everything that did not change stays untouched. Running it twice in a row
 * writes nothing the second time.
 *
 * How a source KPI is found in the Masterfile (first rule that fits wins):
 *   1. same Masterfile KPI name, e.g. "CTB (12 weeks)(ISC)"
 *   2. same KPI name (spaces ignored, "(Function)" suffix ignored) and same definition
 *   3. same definition only (the KPI was renamed in the source -> name is updated)
 *   4. same KPI name only (the definition was changed in the source)
 */

// =====================================================================
// CONFIGURATION
// =====================================================================

/**
 * Function tabs that are not written to 'DMB Masterfile'. The Masterfile has no
 * Commercial Excellence or Finance section; remove a name here to have that
 * tab synced (its section is then created automatically).
 */
const SKIP_FUNCTIONS: string[] = ["Commercial Excellence", "Finance"];

/** Function tab -> 'DMB Masterfile' section title. ISC and Procurement share one section. */
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

/**
 * false = every change found is written to the Masterfile (normal use).
 * true  = nothing is written; the run output only lists what would change.
 */
const PREVIEW_ONLY = false;

/** Delete Masterfile KPIs that were removed from the source. */
const ALLOW_DELETE = true;

/**
 * Safety net: if more than this share of a section's KPIs would be deleted in
 * one run, nothing is deleted in that section (a broken or half-saved source
 * file must not empty the Masterfile). Set to 1 to switch the check off.
 */
const MAX_DELETE_SHARE = 0.5;

const TABLE_STYLE = "TableStyleMedium2";

// =====================================================================
// MASTERFILE LAYOUTS
// =====================================================================

/** Where things sit on one Masterfile tab. Column numbers are 0-based (A = 0). */
interface Layout {
  kind: "DMB" | "MPR";
  sheetName: string;
  columnCount: number;
  nameColumn: number;
  definitionColumn: number;
  unitColumn: number;
  rowTypeColumn: number;
  firstMonthColumn: number;
  dataTypeColumn: number;
  targetLabel: string;
  actualLabel: string;
  headerFirstCells: string[];   // first cell of a section's header row, normalised
  headers: string[];
  sectionOrder: string[];       // used only to place a brand-new section
  nameOverrides: { [definition: string]: string };
}

const DMB_LAYOUT: Layout = {
  kind: "DMB",
  sheetName: "DMB Masterfile",
  columnCount: 22,              // A..V
  nameColumn: 0,                // A  KPI Name
  definitionColumn: 1,          // B  KPI Definition
  unitColumn: 4,                // E  Units
  rowTypeColumn: 7,             // H  Target/ Actual
  firstMonthColumn: 8,          // I  Jan-2026 .. T Dec-2026
  dataTypeColumn: 20,           // U  Data Type
  targetLabel: "Target",
  actualLabel: "Actual",
  headerFirstCells: ["kpi name"],
  headers: [
    "KPI Name", "KPI  Definition", "Operators", "Target AOP 2026", "Units",
    "Metric nature", "Frequency", "Target/ Actual",
    "Jan-2026", "Feb-2026", "Mar-2026", "Apr-2026", "May-2026", "Jun-2026",
    "Jul-2026", "Aug-2026", "Sep-2026", "Oct-2026", "Nov-2026", "Dec-2026",
    "Data Type", "KPI category"
  ],
  sectionOrder: [
    "Quality DMB", "Regulatory", "ISC & Procurement", "R&D", "Marketing",
    "Customer Service", "Commercial Excellence", "NAR", "Europe", "Growth", "Finance"
  ],
  nameOverrides: {
    "scar clsoure % scars closed < 180d": "SCAR Performance (closed less then 180d)",
    "scars on time to plan (ottp)": "SCAR Performance (OTTP)"
  }
};

const MPR_LAYOUT: Layout = {
  kind: "MPR",
  sheetName: "MPR Masterfile",
  columnCount: 23,              // A..W
  nameColumn: 0,                // A  Core KPI's
  definitionColumn: 2,          // C  KPI Definition
  unitColumn: 7,                // H  Units
  rowTypeColumn: 9,             // J  Target/Actual
  firstMonthColumn: 10,         // K  Jan .. V Dec
  dataTypeColumn: 22,           // W  Data Type
  targetLabel: "T",
  actualLabel: "A",
  headerFirstCells: ["core kpi's", "core kpis", "core kpi’s"],
  headers: [
    "Core KPI's", "Metric Owner", "KPI  Definition", "AOP 2026", "Column1", "Column2",
    "Metric nature", "Units", "Frequency", "Target/\nActual",
    "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
    "Data Type"
  ],
  sectionOrder: [
    "Patient Safety & Quality",
    "Customer Focus",
    "Improve Deliverability & Profitability",
    "Roadmap competitiveness & Innovation agility",
    "Drive growth through Commercial Excellence",
    "Environmental, Social & Governance"
  ],
  nameOverrides: {
    "oit growth % (eq)": "OIT Growth % (EQ)",
    "oit growth % (cs)": "OIT Growth % (CS)"
  }
};

// =====================================================================
// TYPES
// =====================================================================

type Cell = string | number | boolean;

/** One row of the JSON returned by extract_functional_dmb.ts / extract_strategic_execution.ts. */
interface SourceRow {
  functionName?: string;         // Functional DMB
  strategicImperative?: string;  // Strategic Execution
  kpiName: string;
  owner?: string;
  definition: string;
  operator?: string;
  target2026: Cell;
  column1?: Cell;
  column2?: Cell;
  unit: string;
  nature: string;
  frequency: string;
  rowType: string;
  jan: Cell; feb: Cell; mar: Cell; apr: Cell; may: Cell; jun: Cell;
  jul: Cell; aug: Cell; sep: Cell; oct: Cell; nov: Cell; dec: Cell;
  dataType?: string;
  category?: string;
  target2026Format?: string;     // number format of target2026 in the source
  formats?: string[];            // number formats of Jan..Dec in the source
}

/** One source KPI with its Target and Actual values. */
interface Entry {
  group: string;          // function tab (DMB) or strategic imperative (MPR)
  section: string;        // Masterfile section title
  kpiName: string;        // name in the source
  displayName: string;    // name written to the Masterfile
  owner: string;
  definition: string;
  operator: string;
  target2026: Cell;
  column1: Cell;
  column2: Cell;
  unit: string;
  nature: string;
  frequency: string;
  dataType: string;
  category: string;
  targetValues: Cell[];
  actualValues: Cell[];
  /** Source number formats, so the Masterfile shows each value as the source does; "" = keep. */
  target2026Format: string;
  targetFormats: string[];
  actualFormats: string[];
}

interface MasterKpi {
  name: string;
  firstRow: number;
  lastRow: number;
  targetRow: number;      // -1 when missing
  actualRow: number;      // -1 when missing
  definitions: string[];  // normalised definitions of its rows
}

interface MasterSection {
  title: string;
  titleRow: number;
  headerRow: number;
  lastRow: number;
  kpis: MasterKpi[];
}

interface Master {
  values: Cell[][];
  formats: string[][];
  sections: MasterSection[];
}

interface Match {
  entry: Entry;
  section: MasterSection;
  kpi: MasterKpi;
  rule: number;
}

interface TabReport {
  sheet: string;
  status: string;
  cellsUpdated: number;
  kpisAdded: number;
  kpisDeleted: number;
  sectionsAdded: number;
  changes: string[];
}

interface SyncReport {
  status: string;
  dryRun: boolean;
  dmb: TabReport;
  mpr: TabReport;
}

// =====================================================================
// MAIN
// =====================================================================

function main(
  workbook: ExcelScript.Workbook,
  functionalJson?: string,
  strategicJson?: string
): SyncReport {
  const isDryRun = PREVIEW_ONLY;

  // Rows are routed by their content, so it does not matter which extract
  // result is passed in which parameter.
  const allRows = parseRows(functionalJson).concat(parseRows(strategicJson));
  const functionalRows = allRows.filter(row => cleanText(row.functionName) !== "");
  const strategicRows = allRows.filter(row => cleanText(row.functionName) === "" &&
    cleanText(row.strategicImperative) !== "");

  const dmb = syncTab(workbook, DMB_LAYOUT, functionalRows, isDryRun);
  const mpr = syncTab(workbook, MPR_LAYOUT, strategicRows, isDryRun);
  const status = `${isDryRun ? "[DRY RUN] " : ""}${dmb.status} | ${mpr.status}`;
  console.log(status);
  return { status: status, dryRun: isDryRun, dmb: dmb, mpr: mpr };
}

/** Rows of one extract result; empty when the text is empty or not JSON. */
function parseRows(json: string | undefined): SourceRow[] {
  if (!json || json.trim().length < 3) return [];
  try {
    const parsed = JSON.parse(json) as SourceRow[];
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    console.log(`Could not read an extract result as JSON: ${String(error)}`);
    return [];
  }
}

/** Sync one Masterfile tab from the rows of its source workbook. */
function syncTab(
  workbook: ExcelScript.Workbook,
  layout: Layout,
  rows: SourceRow[],
  dryRun: boolean
): TabReport {
  const report: TabReport = {
    sheet: layout.sheetName, status: "",
    cellsUpdated: 0, kpisAdded: 0, kpisDeleted: 0, sectionsAdded: 0, changes: []
  };

  if (rows.length === 0) {
    report.status = `${layout.sheetName}: skipped - no rows from its source workbook ` +
      `(check that the extract step runs on that workbook).`;
    return report;
  }
  const entries = buildEntries(layout, rows);
  if (entries.length === 0) {
    report.status = `${layout.sheetName}: skipped (no KPI found in the source JSON).`;
    return report;
  }
  const sheet = workbook.getWorksheet(layout.sheetName);
  if (!sheet) {
    report.status = `${layout.sheetName}: skipped (worksheet not found).`;
    return report;
  }

  // 1. Changed cells of KPIs that already exist.
  let master = readMaster(layout, sheet);
  updateChangedCells(layout, sheet, master, matchEntries(layout, entries, master), report, dryRun);

  // 2. KPIs (and sections) that are new in the source.
  if (dryRun) {
    const matched = matchEntries(layout, entries, master);
    for (const entry of entries) {
      if (!findMatch(matched, entry)) {
        log(report, `ADD    [${entry.section}] ${entry.displayName}`);
        report.kpisAdded++;
      }
    }
  } else {
    addMissingKpis(layout, workbook, sheet, entries, report);
  }

  // 3. KPIs that were removed from the source.
  master = readMaster(layout, sheet);
  deleteRemovedKpis(layout, sheet, master, matchEntries(layout, entries, master), entries, report, dryRun);

  report.status = `${layout.sheetName}: ${report.cellsUpdated} cell(s) updated, ` +
    `${report.kpisAdded} KPI(s) added, ${report.kpisDeleted} KPI(s) deleted, ` +
    `${report.sectionsAdded} section(s) added.`;
  return report;
}

// =====================================================================
// SOURCE -> ENTRIES
// =====================================================================

function buildEntries(layout: Layout, rows: SourceRow[]): Entry[] {
  const entries: Entry[] = [];
  const index: { [key: string]: Entry } = {};
  const skip = SKIP_FUNCTIONS.map(name => name.toLowerCase());
  const isDMB = layout.kind === "DMB";

  const hasTarget: Entry[] = [];

  for (const row of rows) {
    const kpiName = cleanText(row.kpiName);
    const group = isDMB
      ? cleanText(row.functionName) || "General"
      : canonicalImperative(layout, cleanText(row.strategicImperative)) || "Strategic Imperatives";
    if (!kpiName || kpiName.indexOf("#REF") === 0) continue;
    if (isDMB && skip.indexOf(group.toLowerCase()) >= 0) continue;

    const definition = cleanText(row.definition);
    const key = `${group.toLowerCase()}||${kpiName.toLowerCase()}||${definition.toLowerCase()}`;
    const isTarget = isTargetLabel(row.rowType);
    /*
     * A second Target row under the same key is a KPI of its own (e.g. the EQ
     * and CS rows of "Order Intake Growth"); never let it overwrite the first.
     */
    if (index[key] && isTarget && hasTarget.indexOf(index[key]) >= 0) {
      delete index[key];
    }
    if (!index[key]) {
      const unit = cleanText(row.unit);
      index[key] = {
        group: group,
        section: isDMB ? (FUNCTION_TO_SECTION[group.toLowerCase()] || group) : group,
        kpiName: kpiName,
        displayName: kpiName,
        owner: cleanText(row.owner),
        definition: definition,
        operator: cleanText(row.operator),
        target2026: tidy(row.target2026),
        column1: tidy(row.column1),
        column2: tidy(row.column2),
        unit: unit,
        nature: cleanText(row.nature),
        frequency: cleanText(row.frequency),
        dataType: cleanText(row.dataType) || dataTypeForUnit(unit),
        category: cleanText(row.category),
        targetValues: emptyMonths(),
        actualValues: emptyMonths(),
        target2026Format: cleanText(row.target2026Format),
        targetFormats: emptyFormats(),
        actualFormats: emptyFormats()
      };
      entries.push(index[key]);
    }
    const months = [row.jan, row.feb, row.mar, row.apr, row.may, row.jun,
                    row.jul, row.aug, row.sep, row.oct, row.nov, row.dec].map(v => tidy(v));
    const formats = emptyFormats().map((blank, m) => cleanText((row.formats || [])[m]));
    if (isTarget) {
      index[key].targetValues = months;
      index[key].targetFormats = formats;
      hasTarget.push(index[key]);
    } else {
      index[key].actualValues = months;
      index[key].actualFormats = formats;
    }
  }

  // Leftover template rows: no definition and not a single value.
  const populated = entries.filter(entry =>
    entry.definition.length > 0 ||
    entry.targetValues.concat(entry.actualValues).some(v => v !== ""));
  resolveDisplayNames(layout, populated);
  return populated;
}

/**
 * Masterfile KPI names:
 *  - the layout's name overrides first, e.g. "OIT Growth % (EQ)";
 *  - a KPI name used in more than one function tab / imperative gets
 *    "(Group)" added, e.g. "CTB (12 weeks)(ISC)" and "CTB (12 weeks)(Procurement)";
 *  - a name still repeated inside one section is told apart by its definition.
 */
function resolveDisplayNames(layout: Layout, entries: Entry[]): void {
  const groupsByName: { [name: string]: string[] } = {};
  for (const entry of entries) {
    const key = compact(entry.kpiName);
    if (!groupsByName[key]) groupsByName[key] = [];
    if (groupsByName[key].indexOf(entry.group) < 0) groupsByName[key].push(entry.group);
  }
  for (const entry of entries) {
    const override = layout.nameOverrides[normalise(entry.definition)];
    if (override) {
      entry.displayName = override;
    } else if (groupsByName[compact(entry.kpiName)].length > 1) {
      entry.displayName = `${entry.kpiName}(${entry.group})`;
    }
  }
  const bySection: { [key: string]: Entry[] } = {};
  for (const entry of entries) {
    const key = `${sectionKey(entry.section)}||${compact(entry.displayName)}`;
    if (!bySection[key]) bySection[key] = [];
    bySection[key].push(entry);
  }
  for (const key of Object.keys(bySection)) {
    const clashing = bySection[key];
    if (clashing.length < 2) continue;
    clashing.forEach((entry, i) => {
      const tag = cleanText(entry.definition).match(/\(([^()]+)\)\s*$/);
      entry.displayName = `${entry.displayName} (${tag ? tag[1].trim() : String(i + 1)})`;
    });
  }
}

/** The Masterfile cells of the Target or Actual row of a KPI. */
function rowValues(layout: Layout, entry: Entry, isTarget: boolean): Cell[] {
  const label = isTarget ? layout.targetLabel : layout.actualLabel;
  const months = isTarget ? entry.targetValues : entry.actualValues;
  if (layout.kind === "DMB") {
    return ([
      entry.displayName, entry.definition, entry.operator, entry.target2026,
      entry.unit, entry.nature, entry.frequency, label
    ] as Cell[]).concat(months).concat([entry.dataType, entry.category]);
  }
  return ([
    entry.displayName, entry.owner, entry.definition, entry.target2026,
    entry.column1, entry.column2, entry.nature, entry.unit, entry.frequency, label
  ] as Cell[]).concat(months).concat([entry.dataType]);
}

/** Source number formats of the cells of rowValues(); "" = leave the cell's format alone. */
function rowFormats(layout: Layout, entry: Entry, isTarget: boolean): string[] {
  const formats: string[] = [];
  for (let c = 0; c < layout.columnCount; c++) formats.push("");
  formats[3] = entry.target2026Format;            // D: Target AOP 2026 / AOP 2026
  const months = isTarget ? entry.targetFormats : entry.actualFormats;
  for (let m = 0; m < 12; m++) formats[layout.firstMonthColumn + m] = months[m];
  return formats;
}

/**
 * A KPI with Unit "%" and Data Type "Percentage". Not "Production units"
 * (Percentage but "No.") and not NPI On-Time (% kept as Whole Number 90).
 */
function isPercentKpi(unit: string, dataType: string): boolean {
  return cleanText(unit) === "%" && normalise(dataType) === "percentage";
}

/**
 * Plain numbers of a percentage KPI get their % sign: 58.4 -> 0.584 shown as
 * 58.4%, 91 -> 91%, 0.91 -> 91%. Values already shown as % are left alone.
 * Changes the Target AOP 2026 and Jan..Dec cells of `values` / `formats`.
 */
function showAsPercent(layout: Layout, values: Cell[], formats: string[]): void {
  const columns = [3];
  for (let m = 0; m < 12; m++) columns.push(layout.firstMonthColumn + m);
  for (const c of columns) {
    const value = values[c];
    if (typeof value !== "number" || !isFinite(value) || isPercentFormat(formats[c] || "")) continue;
    const shown = Math.abs(value) > 1 ? value : value * 100;
    values[c] = Number((shown / 100).toPrecision(12));
    formats[c] = percentFormatFor(shown);
  }
}

/** "0%", "0.0%" or "0.00%": as many decimals as 58.4 / 91 / 49.25 need. */
function percentFormatFor(shown: number): string {
  for (let places = 0; places < 2; places++) {
    const factor = Math.pow(10, places);
    if (Math.abs(shown * factor - Math.round(shown * factor)) < 0.000001) {
      return places === 0 ? "0%" : "0.0%";
    }
  }
  return "0.00%";
}

/** True for 0% / 0.00% formats; a quoted or escaped "%" is only a literal sign. */
function isPercentFormat(format: string): boolean {
  return format.replace(/"[^"]*"/g, "").replace(/\\./g, "").indexOf("%") >= 0;
}

/**
 * The Masterfile name of a strategic imperative as written in 'AOP Critical',
 * e.g. "Cusotomer Focus" -> "Customer Focus",
 * "Environmental, Social & Governance (ESG)" -> "Environmental, Social & Governance".
 */
function canonicalImperative(layout: Layout, raw: string): string {
  const lower = raw.toLowerCase();
  const pick = (keyword: string): string =>
    layout.sectionOrder.find(name => name.toLowerCase().indexOf(keyword) >= 0) || raw;
  if (!lower) return raw;
  if (lower.indexOf("customer") >= 0 || lower.indexOf("cusotomer") >= 0) return pick("customer");
  if (lower.indexOf("deliverability") >= 0 || lower.indexOf("profitability") >= 0) return pick("deliverability");
  if (lower.indexOf("safety") >= 0 || lower.indexOf("quality") >= 0) return pick("safety");
  if (lower.indexOf("commercial") >= 0 || lower.indexOf("growth") >= 0) return pick("commercial");
  if (lower.indexOf("roadmap") >= 0 || lower.indexOf("innovation") >= 0) return pick("roadmap");
  if (lower.indexOf("esg") >= 0 || lower.indexOf("environmental") >= 0 || lower.indexOf("governance") >= 0) {
    return pick("environmental");
  }
  return raw;
}

/** Every group name a "(Group)" suffix in a Masterfile KPI name can refer to. */
function knownGroups(layout: Layout, entries: Entry[]): string[] {
  const groups = layout.kind === "DMB"
    ? Object.keys(FUNCTION_TO_SECTION)
    : layout.sectionOrder.map(name => name.toLowerCase());
  for (const entry of entries) {
    if (groups.indexOf(entry.group.toLowerCase()) < 0) groups.push(entry.group.toLowerCase());
  }
  return groups;
}

// =====================================================================
// READ A MASTERFILE TAB
// =====================================================================

function readMaster(layout: Layout, sheet: ExcelScript.Worksheet): Master {
  const used = sheet.getUsedRange(true);
  if (!used) return { values: [], formats: [], sections: [] };
  const rowCount = used.getRowIndex() + used.getRowCount();
  const range = sheet.getRangeByIndexes(0, 0, rowCount, layout.columnCount);
  const values = range.getValues() as Cell[][];
  const formats = range.getNumberFormats() as string[][];

  const sections: MasterSection[] = [];
  let current: MasterSection | null = null;

  for (let r = 0; r < values.length; r++) {
    const first = normalise(values[r][layout.nameColumn]);

    if (layout.headerFirstCells.indexOf(first) >= 0) {
      let titleRow = r - 1;
      while (titleRow >= 0 && cleanText(values[titleRow][layout.nameColumn]) === "") titleRow--;
      current = {
        title: titleRow >= 0 ? cleanText(values[titleRow][layout.nameColumn]) : "",
        titleRow: titleRow,
        headerRow: r,
        lastRow: r,
        kpis: []
      };
      sections.push(current);
      continue;
    }

    const rowType = normalise(values[r][layout.rowTypeColumn]);
    if (!current || ["target", "actual", "t", "a"].indexOf(rowType) < 0) continue;

    const isTarget = rowType.charAt(0) === "t";
    const name = cleanText(values[r][layout.nameColumn]);
    const definition = normalise(values[r][layout.definitionColumn]);
    const last = current.kpis.length > 0 ? current.kpis[current.kpis.length - 1] : null;

    // An Actual row directly below its Target row completes the pair.
    if (!isTarget && last && last.actualRow < 0 && last.lastRow === r - 1 &&
        (name === "" || compact(name) === compact(last.name))) {
      last.actualRow = r;
      last.lastRow = r;
      if (last.definitions.indexOf(definition) < 0) last.definitions.push(definition);
    } else {
      current.kpis.push({
        name: name,
        firstRow: r,
        lastRow: r,
        targetRow: isTarget ? r : -1,
        actualRow: isTarget ? -1 : r,
        definitions: [definition]
      });
    }
    current.lastRow = r;
  }

  return { values: values, formats: formats, sections: sections };
}

// =====================================================================
// MATCH SOURCE KPIs TO MASTERFILE KPIs
// =====================================================================

function matchEntries(layout: Layout, entries: Entry[], master: Master): Match[] {
  const matches: Match[] = [];
  const groups = knownGroups(layout, entries);

  for (const section of master.sections) {
    const candidates = entries.filter(e => sectionKey(e.section) === sectionKey(section.title));
    const free = section.kpis.slice();
    const done: Entry[] = [];

    const take = (entry: Entry, kpi: MasterKpi, rule: number): void => {
      matches.push({ entry: entry, section: section, kpi: kpi, rule: rule });
      free.splice(free.indexOf(kpi), 1);
      done.push(entry);
    };
    const open = (): Entry[] => candidates.filter(e => done.indexOf(e) < 0);
    // A master name ending in "(ISC)" only belongs to the ISC tab.
    const sameGroup = (entry: Entry, kpi: MasterKpi): boolean => {
      const suffix = groupSuffix(kpi.name, groups);
      return suffix === "" || suffix === entry.group.toLowerCase();
    };
    const baseName = (kpi: MasterKpi): string => compact(stripGroupSuffix(kpi.name, groups));

    // Rule 1: same Masterfile name.
    for (const entry of open()) {
      const kpi = free.find(k => compact(k.name) === compact(entry.displayName));
      if (kpi) take(entry, kpi, 1);
    }
    // Rule 2: same KPI name and same definition.
    for (const entry of open()) {
      const kpi = free.find(k => baseName(k) === compact(entry.kpiName) &&
        k.definitions.indexOf(normalise(entry.definition)) >= 0 && sameGroup(entry, k));
      if (kpi) take(entry, kpi, 2);
    }
    // Rule 3: same definition, only when that is unambiguous (KPI renamed).
    for (const entry of open()) {
      const definition = normalise(entry.definition);
      if (!definition) continue;
      const hits = free.filter(k => k.definitions.indexOf(definition) >= 0 && sameGroup(entry, k));
      const rivals = open().filter(e => e !== entry && normalise(e.definition) === definition);
      if (hits.length === 1 && rivals.length === 0) take(entry, hits[0], 3);
    }
    // Rule 4: same KPI name, only when unambiguous (definition changed).
    for (const entry of open()) {
      const hits = free.filter(k => baseName(k) === compact(entry.kpiName) && sameGroup(entry, k));
      if (hits.length === 1) take(entry, hits[0], 4);
    }
  }
  return matches;
}

function findMatch(matches: Match[], entry: Entry): Match | undefined {
  return matches.find(m => m.entry === entry);
}

function findSection(master: Master, sectionName: string): MasterSection | undefined {
  return master.sections.find(s => sectionKey(s.title) === sectionKey(sectionName));
}

// =====================================================================
// 1. UPDATE CHANGED CELLS
// =====================================================================

function updateChangedCells(
  layout: Layout,
  sheet: ExcelScript.Worksheet,
  master: Master,
  matches: Match[],
  report: TabReport,
  dryRun: boolean
): void {
  const groups = knownGroups(layout, matches.map(m => m.entry));

  for (const match of matches) {
    const entry = match.entry;
    const kpi = match.kpi;
    const pairs: { row: number; isTarget: boolean }[] = [
      { row: kpi.targetRow, isTarget: true },
      { row: kpi.actualRow, isTarget: false }
    ];

    for (const pair of pairs) {
      if (pair.row < 0) continue;
      const current = master.values[pair.row];
      const currentFormats = master.formats[pair.row] || [];
      const wanted = rowValues(layout, entry, pair.isTarget);
      const wantedFormats = rowFormats(layout, entry, pair.isTarget);

      // Keep the Masterfile name unless the KPI was really renamed in the source.
      const currentName = cleanText(current[layout.nameColumn]);
      const keepName = compact(currentName) === compact(entry.displayName) ||
        compact(stripGroupSuffix(currentName, groups)) === compact(entry.kpiName);
      // Data Type is maintained in the Masterfile; follow the source only when the unit changes.
      const unitChanged = !sameValue(current[layout.unitColumn], entry.unit);
      const dataType = unitChanged ? entry.dataType : (cleanText(current[layout.dataTypeColumn]) || entry.dataType);
      if (isPercentKpi(entry.unit, dataType)) showAsPercent(layout, wanted, wantedFormats);

      for (let c = 0; c < layout.columnCount; c++) {
        if (c === layout.rowTypeColumn) continue;
        if (c === layout.nameColumn && keepName) continue;
        if (c === layout.dataTypeColumn && !unitChanged) continue;

        // Show the value as the source does (100% not 1); a % KPI shows 58.4 as 58.4%.
        const format = wantedFormats[c];
        const fixFormat = format !== "" && format !== (currentFormats[c] || "");
        const where = `[${match.section.title}] ${entry.displayName} ` +
          `${pair.isTarget ? "Target" : "Actual"} ${layout.headers[c].replace(/\n/g, " ")} ` +
          `(${address(pair.row, c)})`;

        if (sameValue(current[c], wanted[c])) {
          if (!fixFormat) continue;
          log(report, `FORMAT ${where}: '${cleanText(current[c])}' format '${currentFormats[c] || ""}' -> '${format}'`);
          report.cellsUpdated++;
          if (!dryRun) sheet.getCell(pair.row, c).setNumberFormat(format);
          continue;
        }

        log(report, `UPDATE ${where}: '${cleanText(current[c])}' -> '${cleanText(wanted[c])}'`);
        report.cellsUpdated++;
        if (!dryRun) {
          const cell = sheet.getCell(pair.row, c);
          if (fixFormat) cell.setNumberFormat(format);
          cell.setValue(wanted[c]);
        }
      }
    }
  }
}

// =====================================================================
// 2. ADD NEW KPIs / SECTIONS
// =====================================================================

function addMissingKpis(
  layout: Layout,
  workbook: ExcelScript.Workbook,
  sheet: ExcelScript.Worksheet,
  entries: Entry[],
  report: TabReport
): void {
  const attempted: Entry[] = [];

  // One insert at a time: every insert shifts the rows below it, so the
  // Masterfile is read again before the next one.
  for (let guard = 0; guard <= entries.length; guard++) {
    const master = readMaster(layout, sheet);
    const matches = matchEntries(layout, entries, master);
    const missing = entries.find(e => !findMatch(matches, e) && attempted.indexOf(e) < 0);
    if (!missing) return;
    attempted.push(missing);

    const section = findSection(master, missing.section);
    if (section) {
      insertKpi(layout, sheet, master, matches, entries, section, missing);
      log(report, `ADD    [${section.title}] ${missing.displayName}`);
    } else {
      addSection(layout, workbook, sheet, master, missing);
      log(report, `ADD    new section '${missing.section}' with ${missing.displayName}`);
      report.sectionsAdded++;
    }
    report.kpisAdded++;
  }
}

/** Insert the Target/Actual pair at the KPI's source position within its section. */
function insertKpi(
  layout: Layout,
  sheet: ExcelScript.Worksheet,
  master: Master,
  matches: Match[],
  entries: Entry[],
  section: MasterSection,
  entry: Entry
): void {
  const siblings = entries.filter(e => sectionKey(e.section) === sectionKey(entry.section));
  const position = siblings.indexOf(entry);

  let insertAt = section.lastRow + 1;
  let anchor: Match | undefined;
  for (let i = position - 1; i >= 0 && !anchor; i--) anchor = findMatch(matches, siblings[i]);
  if (anchor) {
    insertAt = anchor.kpi.lastRow + 1;
  } else {
    for (let i = position + 1; i < siblings.length && !anchor; i++) anchor = findMatch(matches, siblings[i]);
    if (anchor) insertAt = anchor.kpi.firstRow;
  }

  const style = styleSource(master, section, anchor);
  insertRows(sheet, insertAt, 2);
  growTableToRow(sheet, section.headerRow, insertAt + 1);
  writeKpiRows(layout, sheet, insertAt, entry, style, insertAt, 2);
}

/** Add a whole new section (title, header row, table) for a new function tab / imperative. */
function addSection(
  layout: Layout,
  workbook: ExcelScript.Workbook,
  sheet: ExcelScript.Worksheet,
  master: Master,
  entry: Entry
): void {
  const order = layout.sectionOrder.map(name => sectionKey(name));
  const position = order.indexOf(sectionKey(entry.section));
  let next: MasterSection | undefined;
  if (position >= 0) {
    next = master.sections.find(s => order.indexOf(sectionKey(s.title)) > position);
  }

  // Block: [title][header][Target][Actual] + one blank row between sections.
  let start: number;
  let titleRow: number;
  if (next && next.titleRow >= 0) {
    start = next.titleRow;
    titleRow = start;
  } else if (master.sections.length > 0) {
    start = master.sections[master.sections.length - 1].lastRow + 1;
    titleRow = start + 1;
  } else {
    start = 2;
    titleRow = 2;
  }
  const blockSize = 5;
  const headerRow = titleRow + 1;
  const targetRow = headerRow + 1;

  const reference = master.sections.find(s => s.kpis.some(k => k.targetRow >= 0 && k.actualRow >= 0));
  insertRows(sheet, start, blockSize);

  const shift = (row: number): number => (row >= start ? row + blockSize : row);
  if (reference) {
    const pair = reference.kpis.find(k => k.targetRow >= 0 && k.actualRow >= 0) as MasterKpi;
    if (reference.titleRow >= 0) copyRowFormat(layout, sheet, shift(reference.titleRow), titleRow);
    copyRowFormat(layout, sheet, shift(reference.headerRow), headerRow);
    writeKpiRows(layout, sheet, targetRow, entry,
      { targetRow: pair.targetRow, actualRow: pair.actualRow }, start, blockSize);
  } else {
    writeKpiRows(layout, sheet, targetRow, entry, { targetRow: -1, actualRow: -1 }, start, blockSize);
  }

  sheet.getCell(titleRow, layout.nameColumn).setValue(entry.section);
  sheet.getCell(titleRow, layout.nameColumn).getFormat().getFont().setBold(true);
  sheet.getRangeByIndexes(headerRow, 0, 1, layout.columnCount).setValues([layout.headers]);

  const table = sheet.addTable(sheet.getRangeByIndexes(headerRow, 0, 3, layout.columnCount), true);
  table.setPredefinedTableStyle(TABLE_STYLE);
  table.setShowBandedRows(true);
  table.setShowFilterButton(true);
  table.setName(uniqueTableName(workbook, entry.section));
}

/** Rows whose formatting a new Target/Actual pair copies. */
function styleSource(
  master: Master,
  section: MasterSection,
  anchor: Match | undefined
): { targetRow: number; actualRow: number } {
  const complete = (k: MasterKpi): boolean => k.targetRow >= 0 && k.actualRow >= 0;
  let kpi: MasterKpi | undefined = anchor && complete(anchor.kpi) ? anchor.kpi : undefined;
  if (!kpi) kpi = section.kpis.find(k => complete(k));
  if (!kpi) {
    for (const other of master.sections) {
      kpi = other.kpis.find(k => complete(k));
      if (kpi) break;
    }
  }
  return kpi ? { targetRow: kpi.targetRow, actualRow: kpi.actualRow } : { targetRow: -1, actualRow: -1 };
}

/**
 * Write a KPI's two rows at `row`, formatted like the style rows. The style
 * rows were read before `insertedCount` rows were inserted at `insertedAt`.
 */
function writeKpiRows(
  layout: Layout,
  sheet: ExcelScript.Worksheet,
  row: number,
  entry: Entry,
  style: { targetRow: number; actualRow: number },
  insertedAt: number,
  insertedCount: number
): void {
  const shift = (r: number): number => (r >= insertedAt ? r + insertedCount : r);
  if (style.targetRow >= 0) copyRowFormat(layout, sheet, shift(style.targetRow), row);
  if (style.actualRow >= 0) copyRowFormat(layout, sheet, shift(style.actualRow), row + 1);

  const target = rowValues(layout, entry, true);
  const actual = rowValues(layout, entry, false);
  const targetFormats = rowFormats(layout, entry, true);
  const actualFormats = rowFormats(layout, entry, false);
  if (isPercentKpi(entry.unit, entry.dataType)) {
    showAsPercent(layout, target, targetFormats);
    showAsPercent(layout, actual, actualFormats);
  }
  sheet.getRangeByIndexes(row, 0, 2, layout.columnCount).setValues([target, actual]);
  sheet.getRangeByIndexes(row, layout.firstMonthColumn, 1, 12)
    .setNumberFormat(numberFormatFor(entry.targetValues, entry.unit));
  sheet.getRangeByIndexes(row + 1, layout.firstMonthColumn, 1, 12)
    .setNumberFormat(numberFormatFor(entry.actualValues, entry.unit));

  // Cells whose format is known are shown as in the source (or with their % sign).
  [targetFormats, actualFormats].forEach((formats, i) => {
    for (let c = 0; c < layout.columnCount; c++) {
      if (formats[c]) sheet.getCell(row + i, c).setNumberFormat(formats[c]);
    }
  });
}

// =====================================================================
// 3. DELETE REMOVED KPIs
// =====================================================================

function deleteRemovedKpis(
  layout: Layout,
  sheet: ExcelScript.Worksheet,
  master: Master,
  matches: Match[],
  entries: Entry[],
  report: TabReport,
  dryRun: boolean
): void {
  if (!ALLOW_DELETE) return;
  const doomed: MasterKpi[] = [];

  for (const section of master.sections) {
    // A section the source did not deliver at all is left alone.
    if (!entries.some(e => sectionKey(e.section) === sectionKey(section.title))) continue;

    const orphans = section.kpis.filter(k => !matches.some(m => m.kpi === k));
    if (orphans.length === 0) continue;
    if (orphans.length > section.kpis.length * MAX_DELETE_SHARE) {
      log(report, `SKIP   [${section.title}] ${orphans.length} of ${section.kpis.length} KPIs would be ` +
        `deleted - more than ${MAX_DELETE_SHARE * 100}%, nothing deleted in this section.`);
      continue;
    }
    for (const kpi of orphans) {
      log(report, `DELETE [${section.title}] ${kpi.name} (rows ${kpi.firstRow + 1}-${kpi.lastRow + 1})`);
      report.kpisDeleted++;
      doomed.push(kpi);
    }
  }

  if (dryRun) return;
  // Bottom-up, so earlier deletions do not move the rows still to delete.
  doomed.sort((a, b) => b.firstRow - a.firstRow);
  for (const kpi of doomed) {
    sheet.getRangeByIndexes(kpi.firstRow, 0, kpi.lastRow - kpi.firstRow + 1, 1)
      .getEntireRow().delete(ExcelScript.DeleteShiftDirection.up);
  }
}

// =====================================================================
// SHEET HELPERS
// =====================================================================

function insertRows(sheet: ExcelScript.Worksheet, row: number, count: number): void {
  sheet.getRangeByIndexes(row, 0, count, 1).getEntireRow().insert(ExcelScript.InsertShiftDirection.down);
}

function copyRowFormat(layout: Layout, sheet: ExcelScript.Worksheet, fromRow: number, toRow: number): void {
  sheet.getRangeByIndexes(toRow, 0, 1, layout.columnCount)
    .copyFrom(sheet.getRangeByIndexes(fromRow, 0, 1, layout.columnCount), ExcelScript.RangeCopyType.formats);
}

/**
 * Rows inserted inside a table grow it automatically; rows added right below
 * its last row do not, so the table is resized to take them in.
 */
function growTableToRow(sheet: ExcelScript.Worksheet, headerRow: number, lastRow: number): void {
  for (const table of sheet.getTables()) {
    const range = table.getRange();
    const top = range.getRowIndex();
    const bottom = top + range.getRowCount() - 1;
    if (top !== headerRow) continue;
    if (bottom < lastRow) {
      table.resize(sheet.getRangeByIndexes(top, range.getColumnIndex(), lastRow - top + 1, range.getColumnCount()));
    }
    return;
  }
}

function uniqueTableName(workbook: ExcelScript.Workbook, sectionName: string): string {
  let base = sectionName.replace(/[^A-Za-z0-9_]+/g, "_").replace(/^_+|_+$/g, "");
  if (!/^[A-Za-z_]/.test(base)) base = `T_${base}`;
  const taken = workbook.getTables().map(t => t.getName().toLowerCase());
  let name = base;
  for (let i = 2; taken.indexOf(name.toLowerCase()) >= 0; i++) name = `${base}_${i}`;
  return name;
}

function address(row: number, column: number): string {
  let letters = "";
  for (let c = column + 1; c > 0; c = Math.floor((c - 1) / 26)) {
    letters = String.fromCharCode(65 + ((c - 1) % 26)) + letters;
  }
  return `${letters}${row + 1}`;
}

function log(report: TabReport, line: string): void {
  console.log(`${report.sheet}: ${line}`);
  if (report.changes.length < 500) report.changes.push(line);
}

// =====================================================================
// VALUE HELPERS
// =====================================================================

/** Placeholders the source sheets use for "no value" (QTR = quarterly, HY = half-yearly). */
const BLANK_TOKENS = ["", "none", "#ref!", "#n/a", "#div/0!", "#value!", "#name?", "qtr", "hy", "-"];

function cleanText(value: Cell | undefined | null): string {
  if (value === undefined || value === null) return "";
  return String(value).replace(/ /g, " ").trim();
}

/** Lower case, all whitespace folded to one space. */
function normalise(value: Cell | undefined | null): string {
  return cleanText(value).replace(/\s+/g, " ").toLowerCase();
}

/** Lower case, no whitespace at all: "OOH %" and "OOH%" are the same KPI. */
function compact(value: Cell | undefined | null): string {
  return cleanText(value).replace(/\s+/g, "").toLowerCase();
}

/** "Quality DMB", "ISC - DMB" and "Quality" all compare as the section name without "DMB". */
function sectionKey(title: string): string {
  return normalise(title).replace(/\s*-?\s*dmb$/, "").trim();
}

function isTargetLabel(value: Cell): boolean {
  return normalise(value).charAt(0) === "t";
}

function isBlank(value: Cell | undefined | null): boolean {
  if (typeof value === "number") return false;
  return BLANK_TOKENS.indexOf(normalise(value)) >= 0;
}

/** Source value as it should sit in the Masterfile. */
function tidy(value: Cell | undefined | null): Cell {
  if (value === undefined || value === null || isBlank(value)) return "";
  if (typeof value !== "string") return value;
  const text = cleanText(value);
  return /^-?\d+(\.\d+)?$/.test(text) ? Number(text) : text;
}

function asNumber(value: Cell): number | null {
  if (typeof value === "number") return value;
  if (typeof value === "string" && /^-?\d+(\.\d+)?$/.test(cleanText(value))) return Number(cleanText(value));
  return null;
}

/** Equal for the Masterfile: blanks alike, numbers within rounding, text ignoring spacing. */
function sameValue(a: Cell, b: Cell): boolean {
  const left = tidy(a);
  const right = tidy(b);
  if (left === "" || right === "") return left === right;
  const x = asNumber(left);
  const y = asNumber(right);
  if (x !== null && y !== null) return Math.abs(x - y) <= 1e-9 * Math.max(1, Math.abs(x), Math.abs(y));
  return normalise(left) === normalise(right);
}

function emptyMonths(): Cell[] {
  return ["", "", "", "", "", "", "", "", "", "", "", ""];
}

function emptyFormats(): string[] {
  return ["", "", "", "", "", "", "", "", "", "", "", ""];
}

function dataTypeForUnit(unit: string): string {
  const lower = unit.toLowerCase();
  if (lower === "%") return "Percentage";
  if (lower === "mn" || lower === "k" || lower.indexOf("decimal") >= 0) return "Decimal";
  return "Whole Number";
}

/** "(ISC)" at the end of a Masterfile name -> "isc", when ISC is a known group. */
function groupSuffix(name: string, groups: string[]): string {
  const found = cleanText(name).match(/\(([^()]+)\)\s*$/);
  if (!found) return "";
  const suffix = found[1].trim().toLowerCase();
  return groups.indexOf(suffix) >= 0 ? suffix : "";
}

function stripGroupSuffix(name: string, groups: string[]): string {
  return groupSuffix(name, groups) ? cleanText(name).replace(/\s*\([^()]+\)\s*$/, "") : cleanText(name);
}

/**
 * Number format for the month cells of a new row, as in the Masterfile:
 * 0% / 0.0% / 0.00% for percentages, 0 / 0.0 / 0.00 for other numbers.
 * Only a fallback: cells whose source format is known get that format.
 */
function numberFormatFor(values: Cell[], unit: string): string {
  const isPercent = cleanText(unit) === "%";
  let decimals = 0;
  let sawNumber = false;
  for (const value of values) {
    if (typeof value !== "number" || !isFinite(value)) continue;
    sawNumber = true;
    const shown = isPercent ? value * 100 : value;
    let needed = 2;
    for (let places = 0; places <= 2; places++) {
      const factor = Math.pow(10, places);
      if (Math.abs(shown * factor - Math.round(shown * factor)) < 0.000001) {
        needed = places;
        break;
      }
    }
    decimals = Math.max(decimals, needed);
  }
  if (!sawNumber) return isPercent ? "0%" : "General";
  const tail = decimals === 0 ? "" : (decimals === 1 ? ".0" : ".00");
  return isPercent ? `0${tail}%` : `0${tail}`;
}
