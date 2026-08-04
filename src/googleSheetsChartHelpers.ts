// src/googleSheetsChartHelpers.ts
//
// Helpers for reading and modifying embedded charts.
//
// Charts are embedded objects rather than cell values, so none of the value-based
// tools can see them: a sheet holding only charts reads back as empty. Everything
// here works off spreadsheets.get(fields=sheets(properties,charts)) and the
// updateChartSpec / updateEmbeddedObjectPosition batch requests.

import { sheets_v4 } from 'googleapis';
import { UserError } from 'fastmcp';
import { colLettersToIndex, parseA1ToGridRange, parseRange } from './googleSheetsApiHelpers.js';

type Sheets = sheets_v4.Sheets;

export interface SheetIdMaps {
  idByTitle: Map<string, number>;
  titleById: Map<number, string>;
  firstSheetId: number;
}

/**
 * Fetches sheet id/title maps in a single round trip so a tool touching several
 * ranges does not re-request spreadsheet metadata once per range.
 */
export async function getSheetIdMaps(
  sheets: Sheets,
  spreadsheetId: string
): Promise<SheetIdMaps> {
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: 'sheets(properties(sheetId,title))',
  });

  const idByTitle = new Map<string, number>();
  const titleById = new Map<number, string>();
  let firstSheetId: number | null = null;

  for (const sheet of res.data.sheets ?? []) {
    const id = sheet.properties?.sheetId;
    const title = sheet.properties?.title;
    if (id === undefined || id === null || !title) continue;
    idByTitle.set(title, id);
    titleById.set(id, title);
    if (firstSheetId === null) firstSheetId = id;
  }

  if (firstSheetId === null) {
    throw new UserError('Spreadsheet has no sheets.');
  }

  return { idByTitle, titleById, firstSheetId };
}

/** Quotes a sheet name for A1 notation only when it needs it. */
export function quoteSheetName(title: string): string {
  return /^[A-Za-z0-9_]+$/.test(title) ? title : `'${title.replace(/'/g, "''")}'`;
}

/** 0-based column index to letters. 0 -> "A", 26 -> "AA". */
export function colIndexToLetters(col: number): string {
  let letters = '';
  let n = col + 1;
  while (n > 0) {
    n -= 1;
    letters = String.fromCharCode(65 + (n % 26)) + letters;
    n = Math.floor(n / 26);
  }
  return letters;
}

/**
 * Renders a GridRange back into A1 notation, prefixed with its sheet title.
 * Unbounded edges are emitted as open ranges, matching how the Sheets UI shows them.
 */
export function gridRangeToA1(
  range: sheets_v4.Schema$GridRange | undefined | null,
  titleById?: Map<number, string>
): string {
  if (!range) return '';

  const startRow = range.startRowIndex ?? null;
  const endRow = range.endRowIndex ?? null;
  const startCol = range.startColumnIndex ?? null;
  const endCol = range.endColumnIndex ?? null;

  const startCell =
    (startCol !== null ? colIndexToLetters(startCol) : '') + (startRow !== null ? startRow + 1 : '');
  const endCell =
    (endCol !== null ? colIndexToLetters(endCol - 1) : '') + (endRow !== null ? String(endRow) : '');

  const a1 = startCell && startCell === endCell ? startCell : `${startCell}:${endCell}`;

  const sheetId = range.sheetId;
  const title =
    sheetId !== undefined && sheetId !== null ? titleById?.get(sheetId) : undefined;

  return title ? `${quoteSheetName(title)}!${a1}` : a1;
}

/** Resolves an A1 string (optionally "Sheet!A1:B2") against pre-fetched sheet maps. */
export function resolveA1(
  range: string,
  maps: SheetIdMaps,
  defaultSheetId?: number
): sheets_v4.Schema$GridRange {
  const { sheetName, a1Range } = parseRange(range);

  let sheetId: number;
  if (sheetName) {
    const found = maps.idByTitle.get(sheetName);
    if (found === undefined) {
      throw new UserError(
        `Sheet "${sheetName}" not found. Available: ${[...maps.idByTitle.keys()].join(', ')}`
      );
    }
    sheetId = found;
  } else {
    sheetId = defaultSheetId ?? maps.firstSheetId;
  }

  return parseA1ToGridRange(a1Range, sheetId);
}

/** Sheets colors are 0-1 floats; render them as hex for readability. */
export function rgbToHex(color: sheets_v4.Schema$Color | undefined | null): string | undefined {
  if (!color) return undefined;
  const toByte = (v: number | null | undefined) =>
    Math.round(Math.min(1, Math.max(0, v ?? 0)) * 255);
  const hex = [toByte(color.red), toByte(color.green), toByte(color.blue)]
    .map((n) => n.toString(16).padStart(2, '0'))
    .join('');
  return `#${hex.toUpperCase()}`;
}

/** Reads whichever of the deprecated `color` / current `colorStyle` fields is populated. */
export function readColor(
  color: sheets_v4.Schema$Color | undefined | null,
  colorStyle: sheets_v4.Schema$ColorStyle | undefined | null
): string | undefined {
  if (colorStyle?.themeColor) return `theme:${colorStyle.themeColor}`;
  return rgbToHex(colorStyle?.rgbColor ?? color);
}

/**
 * Merges a name cell into a data range when it sits immediately before it, returning
 * a single contiguous GridRange. Returns null when they are not adjacent.
 *
 * Adjacency matters because Google rejects multi-source ChartData in two cases we hit
 * in practice: sources that are not in ascending sheet order, and row-oriented series
 * ("ChartSourceRange ranges require all rows or all columns to have length of 1").
 * One contiguous range sidesteps both.
 */
export function mergeAdjacentNameCell(
  nameCell: sheets_v4.Schema$GridRange,
  dataRange: sheets_v4.Schema$GridRange
): sheets_v4.Schema$GridRange | null {
  if (nameCell.sheetId !== dataRange.sheetId) return null;

  const sameRows =
    nameCell.startRowIndex === dataRange.startRowIndex &&
    nameCell.endRowIndex === dataRange.endRowIndex;
  const sameCols =
    nameCell.startColumnIndex === dataRange.startColumnIndex &&
    nameCell.endColumnIndex === dataRange.endColumnIndex;

  // Label immediately to the left of a row of values.
  if (sameRows && nameCell.endColumnIndex === dataRange.startColumnIndex) {
    return { ...dataRange, startColumnIndex: nameCell.startColumnIndex };
  }

  // Label immediately above a column of values.
  if (sameCols && nameCell.endRowIndex === dataRange.startRowIndex) {
    return { ...dataRange, startRowIndex: nameCell.startRowIndex };
  }

  return null;
}

/** True when a range spans one row and more than one column. */
export function isRowOriented(range: sheets_v4.Schema$GridRange): boolean {
  const rows = (range.endRowIndex ?? 0) - (range.startRowIndex ?? 0);
  const cols = (range.endColumnIndex ?? 0) - (range.startColumnIndex ?? 0);
  return rows === 1 && cols > 1;
}

/**
 * Builds a ChartData source range from an optional name cell plus a data range.
 *
 * The Sheets API has no series-name field. A series is named by the first cell of its
 * source when headerCount is 1. Where the label cell is adjacent to the data it is
 * merged into one range; otherwise it is passed as a second source, which Google only
 * accepts for column-oriented data listed in ascending sheet order.
 */
export function buildChartData(
  dataRanges: sheets_v4.Schema$GridRange[],
  nameCell?: sheets_v4.Schema$GridRange,
  label = 'series'
): sheets_v4.Schema$ChartData {
  if (!nameCell) return { sourceRange: { sources: [...dataRanges] } };

  if (dataRanges.length === 1) {
    const merged = mergeAdjacentNameCell(nameCell, dataRanges[0]);
    if (merged) return { sourceRange: { sources: [merged] } };

    if (isRowOriented(dataRanges[0])) {
      throw new UserError(
        `Google rejects a detached name cell for the row-oriented ${label} ` +
          `${gridRangeToA1(dataRanges[0])}: "ChartSourceRange ranges require all rows or all ` +
          `columns to have length of 1". Put the label in the cell immediately to the left of ` +
          `the values and include it in the range instead, with headerCount 1.`
      );
    }
  }

  return { sourceRange: { sources: [nameCell, ...dataRanges] } };
}

/**
 * Google requires every sourceRange across the domain and all series to appear in
 * ascending sheet order once any of them carries more than one source. Checked here so
 * the caller gets a precise reason instead of the API's generic contiguity complaint.
 */
export function assertSourceOrdering(
  domain: sheets_v4.Schema$ChartData | undefined,
  series: sheets_v4.Schema$BasicChartSeries[]
): void {
  const all: sheets_v4.Schema$ChartData[] = [];
  if (domain) all.push(domain);
  for (const s of series) if (s.series) all.push(s.series);

  const anyMultiSource = all.some((d) => (d.sourceRange?.sources?.length ?? 0) > 1);
  if (!anyMultiSource) return;

  const flat = all.flatMap((d) => d.sourceRange?.sources ?? []);
  for (let i = 1; i < flat.length; i++) {
    const prev = flat[i - 1];
    const cur = flat[i];
    if (prev.sheetId !== cur.sheetId) continue;
    const prevKey = [prev.startRowIndex ?? 0, prev.startColumnIndex ?? 0];
    const curKey = [cur.startRowIndex ?? 0, cur.startColumnIndex ?? 0];
    const ascending =
      curKey[0] > prevKey[0] || (curKey[0] === prevKey[0] && curKey[1] >= prevKey[1]);
    if (!ascending) {
      throw new UserError(
        `Google requires every chart source range to be listed in ascending sheet order once a ` +
          `name cell is used, so ${gridRangeToA1(cur)} cannot follow ${gridRangeToA1(prev)}. ` +
          `Either drop the name cells and reorder freely, or place each label in the cell ` +
          `adjacent to its values so each series is one contiguous range.`
      );
    }
  }
}

/** Flattens a ChartData back to a list of A1 strings for reporting. */
export function chartDataToA1(
  data: sheets_v4.Schema$ChartData | undefined | null,
  titleById: Map<number, string>
): string[] {
  return (data?.sourceRange?.sources ?? []).map((s) => gridRangeToA1(s, titleById));
}

export interface ChartSummary {
  chartId: number;
  anchoredOnSheet: string | null;
  title?: string;
  subtitle?: string;
  chartType: string;
  stackedType?: string;
  legendPosition?: string;
  headerCount?: number;
  axes?: Array<{ position?: string; title?: string }>;
  domains?: string[][];
  series?: Array<{ ranges: string[]; color?: string; targetAxis?: string; type?: string }>;
  pie?: { domain: string[]; series: string[] };
  position?: {
    anchorRow?: number;
    anchorColumn?: number;
    offsetXPixels?: number;
    offsetYPixels?: number;
    widthPixels?: number;
    heightPixels?: number;
  };
  note?: string;
}

/** Turns a raw EmbeddedChart into the readable shape listCharts returns. */
export function summarizeChart(
  chart: sheets_v4.Schema$EmbeddedChart,
  titleById: Map<number, string>
): ChartSummary {
  const spec = chart.spec ?? {};
  const overlay = chart.position?.overlayPosition;
  const anchorSheetId = overlay?.anchorCell?.sheetId;

  const summary: ChartSummary = {
    chartId: chart.chartId ?? -1,
    anchoredOnSheet:
      anchorSheetId !== undefined && anchorSheetId !== null
        ? (titleById.get(anchorSheetId) ?? null)
        : null,
    chartType: 'UNKNOWN',
  };

  if (spec.title) summary.title = spec.title;
  if (spec.subtitle) summary.subtitle = spec.subtitle;

  if (overlay) {
    summary.position = {
      anchorRow: overlay.anchorCell?.rowIndex ?? 0,
      anchorColumn: overlay.anchorCell?.columnIndex ?? 0,
      offsetXPixels: overlay.offsetXPixels ?? 0,
      offsetYPixels: overlay.offsetYPixels ?? 0,
      widthPixels: overlay.widthPixels ?? undefined,
      heightPixels: overlay.heightPixels ?? undefined,
    };
  }

  const basic = spec.basicChart;
  if (basic) {
    summary.chartType = basic.chartType ?? 'BASIC';
    summary.stackedType = basic.stackedType ?? undefined;
    summary.legendPosition = basic.legendPosition ?? undefined;
    summary.headerCount = basic.headerCount ?? 0;
    summary.axes = (basic.axis ?? []).map((a) => ({
      position: a.position ?? undefined,
      title: a.title ?? undefined,
    }));
    summary.domains = (basic.domains ?? []).map((d) => chartDataToA1(d.domain, titleById));
    summary.series = (basic.series ?? []).map((s) => ({
      ranges: chartDataToA1(s.series, titleById),
      color: readColor(s.color, s.colorStyle),
      targetAxis: s.targetAxis ?? undefined,
      type: s.type ?? undefined,
    }));
    summary.note =
      'In a stacked chart the first series is drawn at the bottom. Series names come from the ' +
      'first cell of each series range when headerCount is 1.';
    return summary;
  }

  if (spec.pieChart) {
    summary.chartType = spec.pieChart.pieHole ? 'DONUT' : 'PIE';
    summary.pie = {
      domain: chartDataToA1(spec.pieChart.domain, titleById),
      series: chartDataToA1(spec.pieChart.series, titleById),
    };
    return summary;
  }

  if (spec.treemapChart) {
    summary.chartType = 'TREEMAP';
    return summary;
  }

  for (const key of Object.keys(spec)) {
    if (key.endsWith('Chart')) {
      summary.chartType = key.replace(/Chart$/, '').toUpperCase();
      break;
    }
  }
  summary.note = 'This chart type is reported but not editable through updateChart.';
  return summary;
}

/** Fetches one chart's raw spec and position, or throws if the id is unknown. */
export async function getChartById(
  sheets: Sheets,
  spreadsheetId: string,
  chartId: number
): Promise<{ chart: sheets_v4.Schema$EmbeddedChart; titleById: Map<number, string> }> {
  const res = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: 'sheets(properties(sheetId,title),charts)',
  });

  const titleById = new Map<number, string>();
  for (const sheet of res.data.sheets ?? []) {
    const id = sheet.properties?.sheetId;
    const title = sheet.properties?.title;
    if (id !== undefined && id !== null && title) titleById.set(id, title);
  }

  const known: number[] = [];
  for (const sheet of res.data.sheets ?? []) {
    for (const chart of sheet.charts ?? []) {
      if (chart.chartId === chartId) return { chart, titleById };
      if (chart.chartId !== undefined && chart.chartId !== null) known.push(chart.chartId);
    }
  }

  throw new UserError(
    known.length
      ? `Chart ${chartId} not found. Charts in this spreadsheet: ${known.join(', ')}. Use listCharts.`
      : 'This spreadsheet contains no charts.'
  );
}

/** Validates a 1-based column index used by pie/treemap mapping. */
export function assertColumnLetters(value: string): number {
  if (!/^[A-Za-z]+$/.test(value)) {
    throw new UserError(`Expected column letters like "A" or "AB", got "${value}".`);
  }
  return colLettersToIndex(value);
}
