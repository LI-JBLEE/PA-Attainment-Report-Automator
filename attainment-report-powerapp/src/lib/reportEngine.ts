import JSZip from 'jszip';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

export const REPORT_COLUMNS = [
  'LI_EMP_ID',
  'Person Name',
  'Employee Status',
  'Level Grouping',
  'Level',
  'Fiscal Year',
  'Region',
  'Country',
  'Business_Unit',
  'Measure',
  'Plan_Period',
  'Level_1_Manager',
  'Level_2_Manager',
  'Q1 Credits',
  'Q1 Quota',
  'Q1 Att',
  'Q2 Credits',
  'Q2 Quota',
  'Q2 Att',
  '1H Credits',
  '1H Quota',
  '1H Att',
  'Q3 Credits',
  'Q3 Quota',
  'Q3 Att',
  'Q4 Credits',
  'Q4 Quota',
  'Q4 Att',
  '2H Credits',
  '2H Quota',
  '2H Att',
  'Annual Credits',
  'Annual Quota',
  'Annual Att',
  'Quota Start Date',
  'Quota End Date',
  'Measure Weight',
] as const;

type CellValue = string | number | boolean | Date | null;
type JsonRow = Record<string, unknown>;
type HierarchyItem =
  | { type: 'section'; depth: number; managerName: string }
  | { type: 'row'; depth: number; row: AttainmentRow };

export type AttainmentRow = Record<string, CellValue>;

export interface AttainmentValidationResult {
  rows: AttainmentRow[];
  managerCount: number;
  fiscalYear: string;
  availableRegions: string[];
}

export interface SalesCompValidationResult {
  emailMap: Record<string, string>;
  count: number;
}

export interface ManagerReportRecord {
  id: string;
  fullName: string;
  displayName: string;
  safeName: string;
  region: string;
  email: string | null;
  fileName: string;
  pathInZip: string;
  bytes: Uint8Array;
  attachmentBase64: string;
}

export interface GenerationResult {
  total: number;
  regionCounts: Record<string, number>;
  reports: ManagerReportRecord[];
  fiscalYear: string;
  generatedDate: string;
}

interface BuildIndexResult {
  l1ManagerIds: Set<string>;
  personIdToL1ManagerName: Map<string, string>;
  reportsByManager: Map<string, AttainmentRow[]>;
  managerRegion: Map<string, string>;
  allManagers: string[];
}

const PLAN_PERIOD_OLD_COLUMN = 'Plan_Period;MBO_Description';
const PLAN_PERIOD_COLUMN = 'Plan_Period';
const ATT_SHEET_NAME = 'in';
const SALES_COMP_SHEET_NAME = 'Sheet1';
const SALES_COMP_HEADER_ROW_INDEX = 3;
const SALES_COMP_HEADER_SCAN_LIMIT = 10;
const OTHER_REGION = 'OTHER';
const MAX_EXCEL_OUTLINE_LEVEL = 7;
const ENABLE_OUTLINE_GROUPING = false;
const CELL_ADDRESS_PATTERN = /^[A-Z]+[1-9][0-9]*$/;

const SALES_COMP_REQUIRED_HEADERS = ['Employee ID', 'Email - Work'] as const;

const ATT_COLUMNS = new Set([
  'Q1 Att',
  'Q2 Att',
  '1H Att',
  'Q3 Att',
  'Q4 Att',
  '2H Att',
  'Annual Att',
]);

const NUMBER_COLUMNS = new Set([
  'Q1 Credits',
  'Q1 Quota',
  'Q2 Credits',
  'Q2 Quota',
  '1H Credits',
  '1H Quota',
  'Q3 Credits',
  'Q3 Quota',
  'Q4 Credits',
  'Q4 Quota',
  '2H Credits',
  '2H Quota',
  'Annual Credits',
  'Annual Quota',
]);

const ATTAINMENT_DATE_COLUMNS = new Set(['Person Hire Date', 'Quota Start Date', 'Quota End Date']);
const ATTAINMENT_NUMERIC_COLUMNS = new Set<string>([
  ...Array.from(ATT_COLUMNS),
  ...Array.from(NUMBER_COLUMNS),
  'Measure Weight',
  'Fiscal Year',
]);

const COLUMN_WIDTHS: Record<string, number> = {
  LI_EMP_ID: 12,
  'Person Name': 28,
  'Employee Status': 14,
  'Level Grouping': 14,
  Level: 10,
  'Fiscal Year': 10,
  Region: 10,
  Country: 14,
  Business_Unit: 13,
  Measure: 22,
  Plan_Period: 14,
  Level_1_Manager: 26,
  Level_2_Manager: 26,
  'Q1 Credits': 13,
  'Q1 Quota': 13,
  'Q1 Att': 10,
  'Q2 Credits': 13,
  'Q2 Quota': 13,
  'Q2 Att': 10,
  '1H Credits': 13,
  '1H Quota': 13,
  '1H Att': 10,
  'Q3 Credits': 13,
  'Q3 Quota': 13,
  'Q3 Att': 10,
  'Q4 Credits': 13,
  'Q4 Quota': 13,
  'Q4 Att': 10,
  '2H Credits': 13,
  '2H Quota': 13,
  '2H Att': 10,
  'Annual Credits': 14,
  'Annual Quota': 14,
  'Annual Att': 11,
  'Quota Start Date': 15,
  'Quota End Date': 15,
  'Measure Weight': 14,
};

const LEVEL_FILL_COLORS = ['D6E4F0', 'E8F0E8', 'F5E6F0', 'FFF3E0', 'E0F7FA', 'F1F8E9', 'FCE4EC'];
const SECTION_FILL_COLORS = ['F0E6D3', 'E3D5C1', 'D6C8B3', 'CBBDA8', 'C1B29D', 'B8A894', 'AF9E8B'];

const HEADER_BG_COLOR = '1B3A5C';
const HEADER_FONT_COLOR = 'FFFFFF';
const TITLE_BG_COLOR = '0F2B45';
const SUBTITLE_FONT_COLOR = '666666';
const BORDER_COLOR = 'B0B0B0';

const THIN_BORDER: Partial<ExcelJS.Borders> = {
  top: { style: 'thin', color: { argb: `FF${BORDER_COLOR}` } },
  left: { style: 'thin', color: { argb: `FF${BORDER_COLOR}` } },
  bottom: { style: 'thin', color: { argb: `FF${BORDER_COLOR}` } },
  right: { style: 'thin', color: { argb: `FF${BORDER_COLOR}` } },
};

const HEADER_FONT: Partial<ExcelJS.Font> = {
  name: 'Calibri',
  bold: true,
  size: 10,
  color: { argb: `FF${HEADER_FONT_COLOR}` },
};

const HEADER_FILL: ExcelJS.Fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: `FF${HEADER_BG_COLOR}` },
};

const TITLE_FONT: Partial<ExcelJS.Font> = {
  name: 'Calibri',
  bold: true,
  size: 14,
  color: { argb: `FF${HEADER_FONT_COLOR}` },
};

const TITLE_FILL: ExcelJS.Fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: `FF${TITLE_BG_COLOR}` },
};

const SUBTITLE_FONT: Partial<ExcelJS.Font> = {
  name: 'Calibri',
  size: 10,
  color: { argb: `FF${SUBTITLE_FONT_COLOR}` },
};

const SECTION_FONT: Partial<ExcelJS.Font> = {
  name: 'Calibri',
  bold: true,
  size: 10,
  color: { argb: 'FF333333' },
};

const DATA_FONT: Partial<ExcelJS.Font> = {
  name: 'Calibri',
  size: 10,
};

export function extractManagerName(fullName: string | null | undefined): string {
  if (!fullName) {
    return 'Unknown';
  }

  const clean = fullName.trim();
  const withoutId = clean.replace(/\s*\([^()]*\)\s*$/, '').trim();
  return withoutId || 'Unknown';
}

export function extractManagerId(fullName: string | null | undefined): string | null {
  if (!fullName) {
    return null;
  }

  const match = fullName.match(/\(([^()]+)\)\s*$/);
  return match?.[1]?.trim() || null;
}

export function cleanDisplayName(fullName: string): string {
  return extractManagerName(fullName).replace(/\s*\([^)]*\)/g, '').trim();
}

export function sanitizeFilename(name: string): string {
  const withoutParentheses = name.replace(/\s*\([^)]*\)/g, '');
  return withoutParentheses
    .replace(/[\\/:*?"<>|]/g, '_')
    .trim();
}

export function normalizeEmployeeId(value: unknown): string | null {
  if (value === null || value === undefined) {
    return null;
  }

  const text = String(value).trim();
  if (!text) {
    return null;
  }

  const normalized = text.replace(/\.0+$/, '');
  const withoutLeadingZeros = normalized.replace(/^0+/, '');
  return withoutLeadingZeros || '0';
}

export async function parseAttainmentWorkbook(file: File): Promise<AttainmentValidationResult> {
  const workbook = await readWorkbook(file);
  const sheet = getAttainmentSheet(workbook, file.name);

  if (!sheet) {
    throw new Error(`Cannot find sheet "${ATT_SHEET_NAME}" in attainment workbook.`);
  }

  let rows = parseAttainmentRows(sheet);

  let hasManagerColumn = rows.some((row) => Object.prototype.hasOwnProperty.call(row, 'Level_1_Manager'));
  if (!hasManagerColumn && repairWorksheetRange(sheet)) {
    rows = parseAttainmentRows(sheet);
    hasManagerColumn = rows.some((row) => Object.prototype.hasOwnProperty.call(row, 'Level_1_Manager'));
  }

  if (!hasManagerColumn) {
    throw new Error('Missing required column: Level_1_Manager');
  }

  const managerSet = new Set(
    rows
      .map((row) => asString(row['Level_1_Manager']))
      .filter((value): value is string => Boolean(value)),
  );

  return {
    rows,
    managerCount: managerSet.size,
    fiscalYear: getFiscalYear(rows),
    availableRegions: getAllRegions(rows),
  };
}

export async function parseSalesCompWorkbook(file: File): Promise<SalesCompValidationResult> {
  const workbook = await readWorkbook(file);
  const sheet = workbook.Sheets[SALES_COMP_SHEET_NAME];

  if (!sheet) {
    throw new Error(`Cannot find sheet "${SALES_COMP_SHEET_NAME}" in Sales Compensation workbook.`);
  }

  const headerInfo = findSalesCompHeaderInfo(sheet);
  if (!headerInfo) {
    throw new Error(
      'Cannot find Sales Compensation headers. Expected row 4 to include "Employee ID" and "Email - Work".',
    );
  }

  const emailMap = buildSalesCompEmailMap(sheet, headerInfo);
  if (Object.keys(emailMap).length === 0) {
    throw new Error('Missing valid records for "Employee ID" and "Email - Work".');
  }

  return {
    emailMap,
    count: Object.keys(emailMap).length,
  };
}

export async function generateManagerReports(params: {
  rows: AttainmentRow[];
  emailMap: Record<string, string>;
  selectedRegions: string[];
  fiscalYear: string;
  onProgress?: (current: number, total: number, message: string) => void;
}): Promise<GenerationResult> {
  const { rows, emailMap, selectedRegions, fiscalYear, onProgress } = params;
  if (selectedRegions.length === 0) {
    throw new Error('Select at least one region.');
  }

  const selectedSet = new Set(selectedRegions);
  const index = buildIndexes(rows);
  const reportDate = formatDateCompact(new Date());

  const managersToGenerate = index.allManagers.filter(
    (manager) => selectedSet.has(index.managerRegion.get(manager) || OTHER_REGION),
  );

  const total = managersToGenerate.length;
  const regionCounts: Record<string, number> = {};
  const reports: ManagerReportRecord[] = [];

  for (let i = 0; i < managersToGenerate.length; i += 1) {
    const manager = managersToGenerate[i];
    const hierarchyData = buildHierarchyData(manager, index, 0, new Set<string>());

    if (hierarchyData.length === 0) {
      continue;
    }

    const region = index.managerRegion.get(manager) || OTHER_REGION;
    const cleanName = extractManagerName(manager);
    const safeName = sanitizeFilename(cleanName);
    const fileName = `${fiscalYear}_Attainment_${safeName}_${reportDate}.xlsx`;
    const bytes = await buildWorkbookBytes(manager, hierarchyData, fiscalYear);
    const attachmentBase64 = uint8ArrayToBase64(bytes);
    const managerId = normalizeEmployeeId(extractManagerId(manager));
    const email = managerId ? emailMap[managerId] || null : null;
    const id = `${region}::${safeName}`;

    reports.push({
      id,
      fullName: manager,
      displayName: cleanDisplayName(manager),
      safeName,
      region,
      email,
      fileName,
      pathInZip: `${region}/${fileName}`,
      bytes,
      attachmentBase64,
    });

    regionCounts[region] = (regionCounts[region] || 0) + 1;
    onProgress?.(i + 1, total, `${region}/${fileName}`);

    if ((i + 1) % 4 === 0) {
      await yieldToUi();
    }
  }

  return {
    total: reports.length,
    regionCounts,
    reports,
    fiscalYear,
    generatedDate: formatDateIso(new Date()),
  };
}

export async function buildReportsZip(result: GenerationResult): Promise<Blob> {
  const zip = new JSZip();

  for (const report of result.reports) {
    zip.file(report.pathInZip, report.bytes);
  }

  const metadata = {
    fiscal_year: result.fiscalYear,
    generated_date: result.generatedDate,
    total_reports: result.total,
    managers: result.reports.map((report) => ({
      name: report.displayName,
      safe_name: report.safeName,
      region: report.region,
      email: report.email,
      filepath: report.pathInZip,
    })),
  };

  zip.file('manager_metadata.json', JSON.stringify(metadata, null, 2));
  return zip.generateAsync({ type: 'blob', compression: 'DEFLATE' });
}

function normalizeAttainmentRow(row: JsonRow): AttainmentRow {
  const output: AttainmentRow = {};

  for (const [key, value] of Object.entries(row)) {
    const normalizedKey = key === PLAN_PERIOD_OLD_COLUMN ? PLAN_PERIOD_COLUMN : key;
    output[normalizedKey] = normalizeAttainmentCellValue(normalizedKey, value);
  }

  return output;
}

function normalizeAttainmentCellValue(columnName: string, value: unknown): CellValue {
  if (ATTAINMENT_DATE_COLUMNS.has(columnName)) {
    return normalizeDateCellValue(value);
  }

  if (ATTAINMENT_NUMERIC_COLUMNS.has(columnName)) {
    return normalizeNumericCellValue(value);
  }

  return normalizeCellValue(value);
}

function normalizeCellValue(value: unknown): CellValue {
  if (value === null || value === undefined) {
    return null;
  }
  if (value instanceof Date) {
    return value;
  }
  if (typeof value === 'string') {
    return sanitizeExcelString(value);
  }
  if (typeof value === 'number' || typeof value === 'boolean') {
    return value;
  }
  return null;
}

function normalizeNumericCellValue(value: unknown): CellValue {
  if (value === null || value === undefined) {
    return null;
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return null;
    }

    const parsed = Number(trimmed.replace(/,/g, ''));
    if (Number.isFinite(parsed)) {
      return parsed;
    }

    return sanitizeExcelString(value);
  }

  if (typeof value === 'boolean') {
    return value;
  }

  return null;
}

function normalizeDateCellValue(value: unknown): CellValue {
  if (value === null || value === undefined) {
    return null;
  }

  if (value instanceof Date && Number.isFinite(value.getTime())) {
    return value;
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    return excelSerialToDate(value);
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return null;
    }

    const parsedDate = parseDateStringValue(trimmed);
    if (parsedDate) {
      return parsedDate;
    }

    return sanitizeExcelString(value);
  }

  return null;
}

function asString(value: unknown): string | null {
  if (value === null || value === undefined) {
    return null;
  }

  const text = String(value).trim();
  return text || null;
}

function getFiscalYear(rows: AttainmentRow[]): string {
  const sample = rows
    .map((row) => row['Fiscal Year'])
    .find((value) => value !== null && value !== undefined);

  if (sample === undefined) {
    return 'FY26';
  }

  if (typeof sample === 'number' && Number.isFinite(sample)) {
    return `FY${String(Math.trunc(sample)).slice(-2).padStart(2, '0')}`;
  }

  const asText = String(sample);
  const fourDigit = asText.match(/(20\d{2})/);
  if (fourDigit?.[1]) {
    return `FY${fourDigit[1].slice(-2)}`;
  }

  const fyPattern = asText.match(/FY\s*(\d{2})/i);
  if (fyPattern?.[1]) {
    return `FY${fyPattern[1]}`;
  }

  return 'FY26';
}

function getAttainmentSheet(workbook: XLSX.WorkBook, fileName: string): XLSX.WorkSheet | null {
  const defaultSheet = workbook.Sheets[ATT_SHEET_NAME];
  if (defaultSheet) {
    return defaultSheet;
  }

  if (!isCsvFileName(fileName)) {
    return null;
  }

  const firstSheetName = workbook.SheetNames[0];
  return firstSheetName ? workbook.Sheets[firstSheetName] || null : null;
}

function isCsvFileName(fileName: string): boolean {
  return fileName.toLowerCase().endsWith('.csv');
}

function getAllRegions(rows: AttainmentRow[]): string[] {
  const index = buildIndexes(rows);
  return Array.from(new Set(index.managerRegion.values())).sort((a, b) => a.localeCompare(b));
}

function buildIndexes(rows: AttainmentRow[]): BuildIndexResult {
  const personIdToL1ManagerName = new Map<string, string>();
  const l1ManagerIds = new Set<string>();
  const reportsByManager = new Map<string, AttainmentRow[]>();
  const managerRegion = buildManagerRegionMap(rows);

  const managerSet = new Set<string>();
  for (const row of rows) {
    const managerName = asString(row['Level_1_Manager']);
    if (!managerName) {
      continue;
    }

    managerSet.add(managerName);
    const managerId = normalizeEmployeeId(extractManagerId(managerName));
    if (managerId) {
      personIdToL1ManagerName.set(managerId, managerName);
      l1ManagerIds.add(managerId);
    }

    const managerRows = reportsByManager.get(managerName) || [];
    managerRows.push(row);
    reportsByManager.set(managerName, managerRows);
  }

  return {
    l1ManagerIds,
    personIdToL1ManagerName,
    reportsByManager,
    managerRegion,
    allManagers: Array.from(managerSet).sort((a, b) => a.localeCompare(b)),
  };
}

function buildManagerRegionMap(rows: AttainmentRow[]): Map<string, string> {
  const managerRegion = new Map<string, string>();
  const idToRegion = new Map<string, string>();
  const reportsByManager = new Map<string, string[]>();

  for (const row of rows) {
    const personName = asString(row['Person Name']);
    const managerName = asString(row['Level_1_Manager']);
    const region = asString(row.Region);

    if (personName && region) {
      const personId = normalizeEmployeeId(extractManagerId(personName));
      if (personId) {
        idToRegion.set(personId, region);
      }
    }

    if (managerName) {
      const existing = reportsByManager.get(managerName) || [];
      if (region) {
        existing.push(region);
      }
      reportsByManager.set(managerName, existing);
    }
  }

  for (const managerName of reportsByManager.keys()) {
    const managerId = normalizeEmployeeId(extractManagerId(managerName));
    if (managerId && idToRegion.has(managerId)) {
      managerRegion.set(managerName, idToRegion.get(managerId) || OTHER_REGION);
      continue;
    }

    const inferred = mode(reportsByManager.get(managerName) || []);
    managerRegion.set(managerName, inferred || OTHER_REGION);
  }

  return managerRegion;
}

function buildHierarchyData(
  managerName: string,
  index: BuildIndexResult,
  depth: number,
  visited: Set<string>,
): HierarchyItem[] {
  if (visited.has(managerName)) {
    return [];
  }
  visited.add(managerName);

  const result: HierarchyItem[] = [];
  const directReports = index.reportsByManager.get(managerName) || [];
  if (directReports.length === 0) {
    return result;
  }

  const directPeople = new Set<string>();
  for (const row of directReports) {
    const personName = asString(row['Person Name']);
    if (personName) {
      directPeople.add(personName);
    }
  }

  const directManagers: string[] = [];
  const directNonManagers: string[] = [];

  for (const personName of directPeople) {
    const personId = normalizeEmployeeId(extractManagerId(personName));
    if (personId && index.l1ManagerIds.has(personId)) {
      directManagers.push(personName);
    } else {
      directNonManagers.push(personName);
    }
  }

  directNonManagers.sort((a, b) => a.localeCompare(b));
  directManagers.sort((a, b) => a.localeCompare(b));

  for (const personName of directNonManagers) {
    for (const row of directReports) {
      if (asString(row['Person Name']) === personName) {
        result.push({ type: 'row', depth, row });
      }
    }
  }

  for (const managerPerson of directManagers) {
    result.push({ type: 'section', depth, managerName: managerPerson });

    for (const row of directReports) {
      if (asString(row['Person Name']) === managerPerson) {
        result.push({ type: 'row', depth, row });
      }
    }

    const personId = normalizeEmployeeId(extractManagerId(managerPerson));
    const l1ManagerName = personId
      ? index.personIdToL1ManagerName.get(personId) || managerPerson
      : managerPerson;

    const subItems = buildHierarchyData(l1ManagerName, index, depth + 1, new Set(visited));
    result.push(...subItems);
  }

  return result;
}

async function buildWorkbookBytes(
  managerName: string,
  hierarchyData: HierarchyItem[],
  fiscalYear: string,
): Promise<Uint8Array> {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Attainment Report');
  const numCols = REPORT_COLUMNS.length;
  const headerRow = 4;
  const cleanName = extractManagerName(managerName).replace(/\s*\([^)]*\)/g, '').trim();
  const dateLong = new Intl.DateTimeFormat('en-US', {
    year: 'numeric',
    month: 'long',
    day: '2-digit',
  }).format(new Date());
  const colLetterLast = XLSX.utils.encode_col(numCols - 1);
  let currentRow = headerRow + 1;

  worksheet.columns = REPORT_COLUMNS.map((columnName) => ({
    key: columnName,
    width: COLUMN_WIDTHS[columnName] || 12,
  }));

  worksheet.properties.outlineLevelRow = 0;
  worksheet.views = [{ state: 'frozen', xSplit: 2, ySplit: headerRow }];
  worksheet.pageSetup = {
    orientation: 'landscape',
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
  };

  worksheet.mergeCells(`A1:${colLetterLast}1`);
  worksheet.getCell('A1').value = `${fiscalYear} Attainment Report — ${cleanName}`;
  worksheet.getCell('A1').font = TITLE_FONT;
  worksheet.getCell('A1').alignment = { horizontal: 'left', vertical: 'middle' };
  worksheet.getRow(1).height = 32;
  applyRowFill(worksheet, 1, numCols, TITLE_FILL);

  worksheet.mergeCells(`A2:${colLetterLast}2`);
  worksheet.getCell('A2').value = `Report Date: ${dateLong}    |    Manager: ${cleanName}`;
  worksheet.getCell('A2').font = SUBTITLE_FONT;
  worksheet.getCell('A2').alignment = { horizontal: 'left', vertical: 'middle' };
  worksheet.getRow(2).height = 22;

  worksheet.getRow(3).height = 6;

  for (let colIndex = 1; colIndex <= numCols; colIndex += 1) {
    const headerCell = worksheet.getCell(headerRow, colIndex);
    headerCell.value = REPORT_COLUMNS[colIndex - 1];
    headerCell.font = HEADER_FONT;
    headerCell.fill = HEADER_FILL;
    headerCell.border = THIN_BORDER;
    headerCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
  }
  worksheet.getRow(headerRow).height = 30;
  worksheet.autoFilter = `A${headerRow}:${colLetterLast}${headerRow}`;

  const groupRanges: Array<{ startRow: number; endRow: number; outlineLevel: number }> = [];
  const sectionStack: Array<{ startRow: number; depth: number }> = [];

  for (const item of hierarchyData) {
    if (item.type === 'section') {
      const sectionDepth = item.depth;

      while (sectionStack.length > 0 && sectionStack[sectionStack.length - 1].depth >= sectionDepth) {
        const closed = sectionStack.pop();
        if (closed && currentRow - 1 > closed.startRow) {
          groupRanges.push({
            startRow: closed.startRow + 1,
            endRow: currentRow - 1,
            outlineLevel: closed.depth + 1,
          });
        }
      }

      const sectionName = extractManagerName(item.managerName);
      const sectionFill = getFillFromColor(SECTION_FILL_COLORS[Math.min(sectionDepth, SECTION_FILL_COLORS.length - 1)]);
      const sectionIndent = '  '.repeat(sectionDepth);

      worksheet.mergeCells(currentRow, 1, currentRow, numCols);
      const sectionCell = worksheet.getCell(currentRow, 1);
      sectionCell.value = `${sectionIndent}▸ Team: ${sectionName}`;
      sectionCell.font = SECTION_FONT;
      sectionCell.alignment = { horizontal: 'left', vertical: 'middle' };

      applyRowFillAndBorder(worksheet, currentRow, numCols, sectionFill, THIN_BORDER);
      worksheet.getRow(currentRow).height = 22;

      sectionStack.push({ startRow: currentRow, depth: sectionDepth });
      currentRow += 1;
      continue;
    }

    const rowFill = getFillFromColor(LEVEL_FILL_COLORS[Math.min(item.depth, LEVEL_FILL_COLORS.length - 1)]);
    const indent = '  '.repeat(item.depth);
    const excelRow = worksheet.getRow(currentRow);

    for (let colIndex = 1; colIndex <= numCols; colIndex += 1) {
      const columnName = REPORT_COLUMNS[colIndex - 1];
      const cell = worksheet.getCell(currentRow, colIndex);
      const rawValue = normalizeExcelCellValue(item.row[columnName] ?? null);

      cell.fill = rowFill;
      cell.border = THIN_BORDER;
      cell.font = DATA_FONT;
      cell.alignment = { vertical: 'middle' };

      if (columnName === 'Person Name' && indent) {
        const nameText = asString(rawValue) || '';
        cell.value = `${indent}${nameText}`;
      } else {
        cell.value = rawValue;
      }

      if (ATT_COLUMNS.has(columnName)) {
        const numeric = asNumber(rawValue);
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
        if (numeric !== null && numeric !== 0) {
          cell.numFmt = '0.0%';
          cell.font = getAttainmentFont(numeric);
        } else {
          cell.font = { name: 'Calibri', size: 10, color: { argb: 'FF999999' } };
        }
      } else if (NUMBER_COLUMNS.has(columnName)) {
        const numeric = asNumber(rawValue);
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
        if (numeric !== null && numeric !== 0) {
          cell.numFmt = '#,##0';
        }
      } else if (columnName === 'Measure Weight') {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        if (asNumber(rawValue) !== null) {
          cell.numFmt = '0%';
        }
      } else if (columnName === 'Quota Start Date' || columnName === 'Quota End Date') {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        const excelDateSerial = toExcelDateSerial(rawValue);
        if (excelDateSerial !== null) {
          cell.value = excelDateSerial;
          cell.numFmt = 'YYYY-MM-DD';
        } else if (rawValue !== null && rawValue !== '') {
          cell.numFmt = 'YYYY-MM-DD';
        }
      } else if (columnName === 'LI_EMP_ID' || columnName === 'Fiscal Year') {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      } else if (
        columnName === 'Employee Status' ||
        columnName === 'Level Grouping' ||
        columnName === 'Level' ||
        columnName === 'Region' ||
        columnName === 'Country' ||
        columnName === 'Business_Unit' ||
        columnName === 'Plan_Period'
      ) {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
      }
    }

    excelRow.height = 18;
    currentRow += 1;
  }

  while (sectionStack.length > 0) {
    const closed = sectionStack.pop();
    if (closed && currentRow - 1 > closed.startRow) {
      groupRanges.push({
        startRow: closed.startRow + 1,
        endRow: currentRow - 1,
        outlineLevel: closed.depth + 1,
      });
    }
  }

  if (ENABLE_OUTLINE_GROUPING) {
    groupRanges
      .sort((a, b) => a.outlineLevel - b.outlineLevel)
      .forEach((groupRange) => {
        const outlineLevel = Math.min(groupRange.outlineLevel, MAX_EXCEL_OUTLINE_LEVEL);
        for (let rowIndex = groupRange.startRow; rowIndex <= groupRange.endRow; rowIndex += 1) {
          worksheet.getRow(rowIndex).outlineLevel = outlineLevel;
        }
      });
  }

  const maxOutlineLevel = ENABLE_OUTLINE_GROUPING
    ? groupRanges.reduce((maxValue, groupRange) => {
      const current = Math.min(groupRange.outlineLevel, MAX_EXCEL_OUTLINE_LEVEL);
      return Math.max(maxValue, current);
    }, 0)
    : 0;

  if (maxOutlineLevel > 0) {
    worksheet.properties.outlineLevelRow = maxOutlineLevel;
    worksheet.properties.outlineProperties = { summaryBelow: false, summaryRight: true };
  }

  const buffer = await workbook.xlsx.writeBuffer();
  return sanitizeWorkbookXml(new Uint8Array(buffer), maxOutlineLevel);
}

function normalizeExcelCellValue(value: CellValue): string | number | boolean | Date | null {
  if (value === undefined || value === null) {
    return null;
  }
  if (typeof value === 'string') {
    return sanitizeExcelString(value);
  }
  return value;
}

function sanitizeExcelString(value: string): string {
  const normalizedLineBreak = value.replace(/\r\n/g, '\n');
  let cleaned = '';

  for (let i = 0; i < normalizedLineBreak.length; i += 1) {
    const charCode = normalizedLineBreak.charCodeAt(i);
    const isDisallowedControl = charCode < 32 && charCode !== 9 && charCode !== 10 && charCode !== 13;
    if (!isDisallowedControl) {
      cleaned += normalizedLineBreak[i];
    }
  }

  if (cleaned.length > 32767) {
    return cleaned.slice(0, 32767);
  }
  return cleaned;
}

function applyRowFillAndBorder(
  worksheet: ExcelJS.Worksheet,
  rowNumber: number,
  numCols: number,
  fill: ExcelJS.Fill,
  border: Partial<ExcelJS.Borders>,
): void {
  for (let colIndex = 1; colIndex <= numCols; colIndex += 1) {
    const cell = worksheet.getCell(rowNumber, colIndex);
    cell.fill = fill;
    cell.border = border;
  }
}

function applyRowFill(worksheet: ExcelJS.Worksheet, rowNumber: number, numCols: number, fill: ExcelJS.Fill): void {
  for (let colIndex = 1; colIndex <= numCols; colIndex += 1) {
    const cell = worksheet.getCell(rowNumber, colIndex);
    cell.fill = fill;
  }
}

function getFillFromColor(colorHex: string): ExcelJS.Fill {
  return {
    type: 'pattern',
    pattern: 'solid',
    fgColor: { argb: `FF${colorHex}` },
  };
}

function getAttainmentFont(value: number): Partial<ExcelJS.Font> {
  if (value >= 1.0) {
    return { name: 'Calibri', size: 10, bold: true, color: { argb: 'FF27AE60' } };
  }
  if (value >= 0.8) {
    return { name: 'Calibri', size: 10, color: { argb: 'FFF39C12' } };
  }
  return { name: 'Calibri', size: 10, color: { argb: 'FFE74C3C' } };
}

function asNumber(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return null;
    }

    const parsed = Number(trimmed.replace(/,/g, ''));
    return Number.isFinite(parsed) ? parsed : null;
  }
  return null;
}

function toExcelDateSerial(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return Math.round(value);
  }

  if (value instanceof Date && Number.isFinite(value.getTime())) {
    return dateToExcelSerial(value);
  }

  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (!trimmed) {
      return null;
    }

    const parsedDate = parseDateStringValue(trimmed);
    if (parsedDate) {
      return dateToExcelSerial(parsedDate);
    }
  }

  return null;
}

function excelSerialToDate(value: number): Date | null {
  if (!Number.isFinite(value)) {
    return null;
  }

  const utcMillis = Math.round((value - 25569) * 86400000);
  const date = new Date(utcMillis);
  return Number.isFinite(date.getTime()) ? date : null;
}

function dateToExcelSerial(value: Date): number {
  return Math.round(value.getTime() / 86400000 + 25569);
}

function parseDateStringValue(value: string): Date | null {
  const isoMatch = value.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (isoMatch) {
    return createUtcDate(Number(isoMatch[1]), Number(isoMatch[2]), Number(isoMatch[3]));
  }

  const slashMatch = value.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    return createUtcDate(Number(slashMatch[3]), Number(slashMatch[1]), Number(slashMatch[2]));
  }

  return null;
}

function createUtcDate(year: number, month: number, day: number): Date | null {
  if (
    !Number.isInteger(year) ||
    !Number.isInteger(month) ||
    !Number.isInteger(day) ||
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > 31
  ) {
    return null;
  }

  const date = new Date(Date.UTC(year, month - 1, day));
  if (
    !Number.isFinite(date.getTime()) ||
    date.getUTCFullYear() !== year ||
    date.getUTCMonth() !== month - 1 ||
    date.getUTCDate() !== day
  ) {
    return null;
  }

  return date;
}

async function sanitizeWorkbookXml(bytes: Uint8Array, maxOutlineLevel: number): Promise<Uint8Array> {
  const zip = await JSZip.loadAsync(bytes);
  const worksheetPath = 'xl/worksheets/sheet1.xml';
  const sheetEntry = zip.file(worksheetPath);
  if (!sheetEntry) {
    return bytes;
  }

  const sourceXml = await sheetEntry.async('string');
  let normalizedXml = sourceXml
    .replace(/\scollapsed="1"/g, '')
    .replace(/outlineLevelRow="\d+"/g, `outlineLevelRow="${maxOutlineLevel}"`);

  if (maxOutlineLevel === 0) {
    normalizedXml = normalizedXml.replace(/<outlinePr[^>]*\/>/g, '');
  }

  zip.file(worksheetPath, normalizedXml);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

function mode(values: string[]): string | null {
  if (values.length === 0) {
    return null;
  }

  const counts = new Map<string, number>();
  for (const value of values) {
    counts.set(value, (counts.get(value) || 0) + 1);
  }

  let maxValue: string | null = null;
  let maxCount = -1;
  for (const [value, count] of counts.entries()) {
    if (count > maxCount) {
      maxCount = count;
      maxValue = value;
    }
  }
  return maxValue;
}

function uint8ArrayToBase64(bytes: Uint8Array): string {
  const alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  let output = '';
  let i = 0;

  while (i < bytes.length) {
    const byte1 = bytes[i];
    const byte2 = i + 1 < bytes.length ? bytes[i + 1] : NaN;
    const byte3 = i + 2 < bytes.length ? bytes[i + 2] : NaN;

    const enc1 = byte1 >> 2;
    const enc2 = ((byte1 & 0x03) << 4) | ((Number.isNaN(byte2) ? 0 : byte2) >> 4);
    const enc3 = Number.isNaN(byte2) ? 64 : (((byte2 & 0x0f) << 2) | ((Number.isNaN(byte3) ? 0 : byte3) >> 6));
    const enc4 = Number.isNaN(byte3) ? 64 : (byte3 & 0x3f);

    output += alphabet[enc1];
    output += alphabet[enc2];
    output += enc3 === 64 ? '=' : alphabet[enc3];
    output += enc4 === 64 ? '=' : alphabet[enc4];

    i += 3;
  }

  return output;
}

function formatDateCompact(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}${month}${day}`;
}

function formatDateIso(date: Date): string {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function parseAttainmentRows(sheet: XLSX.WorkSheet): AttainmentRow[] {
  return XLSX.utils.sheet_to_json<JsonRow>(sheet, { defval: null, raw: true }).map(normalizeAttainmentRow);
}

function findSalesCompHeaderInfo(
  sheet: XLSX.WorkSheet,
): { rowIndex: number; employeeIdCol: number; emailCol: number } | null {
  const rowValues = new Map<number, Map<number, string>>();

  for (const key of Object.keys(sheet)) {
    if (!CELL_ADDRESS_PATTERN.test(key)) {
      continue;
    }

    const cellAddress = XLSX.utils.decode_cell(key);
    if (cellAddress.r > SALES_COMP_HEADER_SCAN_LIMIT) {
      continue;
    }

    const value = asString(readWorksheetCellValue(sheet[key]));
    if (!value) {
      continue;
    }

    let row = rowValues.get(cellAddress.r);
    if (!row) {
      row = new Map<number, string>();
      rowValues.set(cellAddress.r, row);
    }
    row.set(cellAddress.c, value);
  }

  const candidateRowIndexes = [
    SALES_COMP_HEADER_ROW_INDEX,
    ...Array.from({ length: SALES_COMP_HEADER_SCAN_LIMIT + 1 }, (_, index) => index).filter(
      (index) => index !== SALES_COMP_HEADER_ROW_INDEX,
    ),
  ];

  for (const rowIndex of candidateRowIndexes) {
    const row = rowValues.get(rowIndex);
    if (!row) {
      continue;
    }

    let employeeIdCol: number | null = null;
    let emailCol: number | null = null;

    for (const [columnIndex, headerText] of row.entries()) {
      if (headerText === SALES_COMP_REQUIRED_HEADERS[0]) {
        employeeIdCol = columnIndex;
      }
      if (headerText === SALES_COMP_REQUIRED_HEADERS[1]) {
        emailCol = columnIndex;
      }
    }

    if (employeeIdCol !== null && emailCol !== null) {
      return {
        rowIndex,
        employeeIdCol,
        emailCol,
      };
    }
  }

  return null;
}

function buildSalesCompEmailMap(
  sheet: XLSX.WorkSheet,
  headerInfo: { rowIndex: number; employeeIdCol: number; emailCol: number },
): Record<string, string> {
  const emailMap: Record<string, string> = {};
  const employeeIdsByRow = new Map<number, unknown>();
  const emailsByRow = new Map<number, unknown>();

  for (const key of Object.keys(sheet)) {
    if (!CELL_ADDRESS_PATTERN.test(key)) {
      continue;
    }

    const cellAddress = XLSX.utils.decode_cell(key);
    if (cellAddress.r <= headerInfo.rowIndex) {
      continue;
    }

    const cellValue = readWorksheetCellValue(sheet[key]);
    if (cellAddress.c === headerInfo.employeeIdCol) {
      employeeIdsByRow.set(cellAddress.r, cellValue);
    } else if (cellAddress.c === headerInfo.emailCol) {
      emailsByRow.set(cellAddress.r, cellValue);
    }
  }

  for (const [rowIndex, rawEmployeeId] of employeeIdsByRow.entries()) {
    const employeeId = normalizeEmployeeId(rawEmployeeId);
    const email = asString(emailsByRow.get(rowIndex));
    if (!employeeId || !email) {
      continue;
    }
    emailMap[employeeId] = email;
  }

  return emailMap;
}

function readWorksheetCellValue(cell: XLSX.CellObject | undefined): unknown {
  if (!cell) {
    return null;
  }

  if ('v' in cell && cell.v !== undefined) {
    return cell.v;
  }

  if ('w' in cell && cell.w !== undefined) {
    return cell.w;
  }

  return null;
}

function repairWorksheetRange(sheet: XLSX.WorkSheet): boolean {
  let minRow = Number.POSITIVE_INFINITY;
  let minCol = Number.POSITIVE_INFINITY;
  let maxRow = -1;
  let maxCol = -1;

  for (const key of Object.keys(sheet)) {
    if (!CELL_ADDRESS_PATTERN.test(key)) {
      continue;
    }

    const cell = XLSX.utils.decode_cell(key);
    minRow = Math.min(minRow, cell.r);
    minCol = Math.min(minCol, cell.c);
    maxRow = Math.max(maxRow, cell.r);
    maxCol = Math.max(maxCol, cell.c);
  }

  if (maxRow < 0 || maxCol < 0) {
    return false;
  }

  const recoveredRange = XLSX.utils.encode_range({
    s: { r: minRow, c: minCol },
    e: { r: maxRow, c: maxCol },
  });

  const currentRangeRef = sheet['!ref'];
  if (!currentRangeRef) {
    sheet['!ref'] = recoveredRange;
    return true;
  }

  try {
    const currentRange = XLSX.utils.decode_range(currentRangeRef);
    const recoveredBounds = XLSX.utils.decode_range(recoveredRange);
    const needsExpansion =
      recoveredBounds.s.r < currentRange.s.r ||
      recoveredBounds.s.c < currentRange.s.c ||
      recoveredBounds.e.r > currentRange.e.r ||
      recoveredBounds.e.c > currentRange.e.c;

    if (!needsExpansion) {
      return false;
    }
  } catch {
    sheet['!ref'] = recoveredRange;
    return true;
  }

  // Some SCR exports carry a truncated !ref (for example, "A1") even though the cells exist.
  sheet['!ref'] = recoveredRange;
  return true;
}

async function readWorkbook(file: File): Promise<XLSX.WorkBook> {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: 'array', cellDates: true });
}

function yieldToUi(): Promise<void> {
  return new Promise((resolve) => {
    window.setTimeout(resolve, 0);
  });
}
