import type { FastMCP } from 'fastmcp';
import { UserError } from 'fastmcp';
import { z } from 'zod';
import { sheets_v4 } from 'googleapis';
import { getSheetsClient } from '../../clients.js';
import { hexToRgb } from '../../googleSheetsApiHelpers.js';
import {
  assertSourceOrdering,
  buildChartData,
  getChartById,
  getSheetIdMaps,
  resolveA1,
  summarizeChart,
} from '../../googleSheetsChartHelpers.js';

const seriesSchema = z.object({
  range: z
    .string()
    .describe(
      'A1 range holding this series\' values, e.g. "Schedule!B10:BZ10" for a row or "Data!C2:C80" for a column.'
    ),
  nameCell: z
    .string()
    .optional()
    .describe(
      'Single cell holding this series\' legend name, e.g. "Data!B1". Merged into the range when it sits ' +
        'immediately before the values (left of a row, above a column), which is the only form Google ' +
        'accepts for row-oriented series. Sets headerCount to 1 automatically.'
    ),
  color: z
    .string()
    .optional()
    .describe('Explicit series color as hex, e.g. "#DB4437". Omit to leave the existing color.'),
  targetAxis: z
    .enum(['LEFT_AXIS', 'RIGHT_AXIS'])
    .optional()
    .describe('Which value axis this series plots against. Defaults to LEFT_AXIS.'),
});

export function register(server: FastMCP) {
  server.addTool({
    name: 'updateChart',
    description:
      'Updates an existing chart in place, keeping its ID and every property you do not pass. ' +
      'Can change the title, chart type, stacking, legend position, axis titles, the domain (x-axis) range, ' +
      'the full series list including stacking order and per-series colors and names, and the chart\'s ' +
      'anchor cell and pixel size. Get the chart ID from listCharts. ' +
      'Stacking order follows the series array: index 0 is drawn at the bottom. ' +
      'The Sheets API has no series-name field — a series is named by the first cell of its source range when ' +
      'headerCount is 1, which is what nameCell supplies. For that to work the label must sit immediately ' +
      'before the values (left of a row, above a column) so the two merge into one contiguous range; Google ' +
      'rejects a detached label on a row-oriented series outright. If the label lives somewhere else, either ' +
      'move it next to the data or omit nameCell and let the series go unnamed. ' +
      'Not settable through the API: the value axis number format. It follows the source cells, so format ' +
      'those as currency to get a currency axis, and make sure no stray date-formatted cell sits in the range.',
    parameters: z.object({
      spreadsheetId: z
        .string()
        .describe(
          'The spreadsheet ID — the long string between /d/ and /edit in a Google Sheets URL.'
        ),
      chartId: z.number().int().describe('Numeric chart ID from listCharts.'),

      title: z.string().optional().describe('Chart title. Pass an empty string to clear it.'),
      subtitle: z.string().optional().describe('Chart subtitle. Pass an empty string to clear it.'),
      chartType: z
        .enum(['BAR', 'COLUMN', 'LINE', 'AREA', 'SCATTER', 'STEPPED_AREA', 'COMBO'])
        .optional()
        .describe('Basic chart type. Omit to keep the current type.'),
      stackedType: z
        .enum(['NOT_STACKED', 'STACKED', 'PERCENT_STACKED'])
        .optional()
        .describe('Stacking mode for bar/column/area charts.'),
      legendPosition: z
        .enum([
          'BOTTOM_LEGEND',
          'LEFT_LEGEND',
          'RIGHT_LEGEND',
          'TOP_LEGEND',
          'NO_LEGEND',
          'LABELED_LEGEND',
        ])
        .optional()
        .describe('Where the legend is drawn.'),

      bottomAxisTitle: z.string().optional().describe('Title for the horizontal axis.'),
      leftAxisTitle: z.string().optional().describe('Title for the left value axis.'),
      rightAxisTitle: z.string().optional().describe('Title for the right value axis.'),

      domainRange: z
        .string()
        .optional()
        .describe(
          'A1 range holding the x-axis labels, e.g. "Schedule!B1:BZ1". Replaces the current domain.'
        ),
      domainNameCell: z
        .string()
        .optional()
        .describe(
          'Single cell prepended to the domain so headerCount consumes it instead of a real label. ' +
            'Supply this whenever series use nameCell and the domain runs in the same orientation.'
        ),

      series: z
        .array(seriesSchema)
        .min(1)
        .optional()
        .describe(
          'Replaces the entire series list, in stacking order (index 0 = bottom). Omit to leave series untouched.'
        ),
      headerCount: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe(
          'How many leading rows/columns of each source are labels rather than data. Set automatically to 1 ' +
            'when any series supplies a nameCell.'
        ),

      anchorRow: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe('0-based row of the anchor cell. Moves the chart.'),
      anchorColumn: z
        .number()
        .int()
        .min(0)
        .optional()
        .describe('0-based column of the anchor cell. Moves the chart.'),
      offsetXPixels: z.number().int().optional().describe('Horizontal offset from the anchor cell.'),
      offsetYPixels: z.number().int().optional().describe('Vertical offset from the anchor cell.'),
      widthPixels: z.number().int().min(1).optional().describe('Chart width in pixels.'),
      heightPixels: z.number().int().min(1).optional().describe('Chart height in pixels.'),
    }),
    execute: async (args, { log }) => {
      const sheets = await getSheetsClient();
      log.info(`Updating chart ${args.chartId} in spreadsheet ${args.spreadsheetId}`);

      try {
        const { chart } = await getChartById(sheets, args.spreadsheetId, args.chartId);

        // updateChartSpec replaces the whole spec, so start from the existing one
        // and patch it. Anything the caller does not pass survives untouched.
        const spec: sheets_v4.Schema$ChartSpec = JSON.parse(JSON.stringify(chart.spec ?? {}));

        const wantsSpecChange =
          args.title !== undefined ||
          args.subtitle !== undefined ||
          args.chartType !== undefined ||
          args.stackedType !== undefined ||
          args.legendPosition !== undefined ||
          args.bottomAxisTitle !== undefined ||
          args.leftAxisTitle !== undefined ||
          args.rightAxisTitle !== undefined ||
          args.domainRange !== undefined ||
          args.series !== undefined ||
          args.headerCount !== undefined;

        const wantsPositionChange =
          args.anchorRow !== undefined ||
          args.anchorColumn !== undefined ||
          args.offsetXPixels !== undefined ||
          args.offsetYPixels !== undefined ||
          args.widthPixels !== undefined ||
          args.heightPixels !== undefined;

        if (!wantsSpecChange && !wantsPositionChange) {
          throw new UserError('Nothing to update: pass at least one property to change.');
        }

        const requests: sheets_v4.Schema$Request[] = [];

        if (wantsSpecChange) {
          if (args.title !== undefined) spec.title = args.title;
          if (args.subtitle !== undefined) spec.subtitle = args.subtitle;

          const basic = spec.basicChart;
          if (!basic) {
            throw new UserError(
              'This chart is not a basic (bar/column/line/area/scatter) chart, so its spec cannot be edited here. ' +
                'Only the title and position are editable for pie, donut and treemap charts.'
            );
          }

          if (args.chartType !== undefined) basic.chartType = args.chartType;
          if (args.stackedType !== undefined) basic.stackedType = args.stackedType;
          if (args.legendPosition !== undefined) basic.legendPosition = args.legendPosition;

          const axisTitles: Array<[string, string | undefined]> = [
            ['BOTTOM_AXIS', args.bottomAxisTitle],
            ['LEFT_AXIS', args.leftAxisTitle],
            ['RIGHT_AXIS', args.rightAxisTitle],
          ];
          for (const [position, title] of axisTitles) {
            if (title === undefined) continue;
            basic.axis = basic.axis ?? [];
            const existing = basic.axis.find((a) => a.position === position);
            if (existing) existing.title = title;
            else basic.axis.push({ position, title });
          }

          const needsResolution =
            args.domainRange !== undefined || args.series !== undefined;
          const maps = needsResolution
            ? await getSheetIdMaps(sheets, args.spreadsheetId)
            : null;

          if (args.domainRange !== undefined && maps) {
            const dataRange = resolveA1(args.domainRange, maps);
            const nameCell = args.domainNameCell
              ? resolveA1(args.domainNameCell, maps)
              : undefined;
            basic.domains = [
              {
                domain: buildChartData([dataRange], nameCell, 'domain'),
                reversed: basic.domains?.[0]?.reversed ?? false,
              },
            ];
          }

          if (args.series !== undefined && maps) {
            const previous = basic.series ?? [];
            basic.series = args.series.map((s, i) => {
              const dataRange = resolveA1(s.range, maps);
              const nameCell = s.nameCell ? resolveA1(s.nameCell, maps) : undefined;

              // Carry forward styling from the series that occupied this slot so an
              // order swap does not silently drop point styles or dash types.
              const carried: sheets_v4.Schema$BasicChartSeries = previous[i]
                ? { ...previous[i] }
                : {};
              delete carried.series;

              const built: sheets_v4.Schema$BasicChartSeries = {
                ...carried,
                series: buildChartData([dataRange], nameCell, `series ${i + 1}`),
                targetAxis: s.targetAxis ?? carried.targetAxis ?? 'LEFT_AXIS',
              };

              if (s.color !== undefined) {
                const rgb = hexToRgb(s.color);
                if (!rgb) {
                  throw new UserError(
                    `Invalid color "${s.color}" for series ${i + 1}. Use hex like "#DB4437".`
                  );
                }
                built.color = rgb;
                built.colorStyle = { rgbColor: rgb };
              }

              return built;
            });

            if (args.series.some((s) => s.nameCell) && args.headerCount === undefined) {
              basic.headerCount = 1;
            }
          }

          if (args.headerCount !== undefined) basic.headerCount = args.headerCount;

          assertSourceOrdering(basic.domains?.[0]?.domain ?? undefined, basic.series ?? []);

          requests.push({
            updateChartSpec: {
              chartId: args.chartId,
              spec,
            },
          });
        }

        if (wantsPositionChange) {
          const overlay = chart.position?.overlayPosition;
          if (!overlay) {
            throw new UserError(
              'This chart occupies its own sheet rather than floating over cells, so it cannot be moved or resized.'
            );
          }

          const anchor = overlay.anchorCell ?? {};
          requests.push({
            updateEmbeddedObjectPosition: {
              objectId: args.chartId,
              newPosition: {
                overlayPosition: {
                  anchorCell: {
                    sheetId: anchor.sheetId,
                    rowIndex: args.anchorRow ?? anchor.rowIndex ?? 0,
                    columnIndex: args.anchorColumn ?? anchor.columnIndex ?? 0,
                  },
                  offsetXPixels: args.offsetXPixels ?? overlay.offsetXPixels ?? 0,
                  offsetYPixels: args.offsetYPixels ?? overlay.offsetYPixels ?? 0,
                  widthPixels: args.widthPixels ?? overlay.widthPixels ?? undefined,
                  heightPixels: args.heightPixels ?? overlay.heightPixels ?? undefined,
                },
              },
              // The API rejects sub-field masks here ("Invalid field: overlay_position"), so the
              // whole position is replaced. Every value above is merged from the existing overlay,
              // which keeps unspecified properties intact.
              fields: '*',
            },
          });
        }

        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: args.spreadsheetId,
          requestBody: { requests },
        });

        const { chart: updated, titleById } = await getChartById(
          sheets,
          args.spreadsheetId,
          args.chartId
        );

        return JSON.stringify(
          {
            updated: true,
            chart: summarizeChart(updated, titleById),
          },
          null,
          2
        );
      } catch (error: any) {
        log.error(`Error updating chart: ${error.message || error}`);
        if (error instanceof UserError) throw error;
        throw new UserError(`Failed to update chart: ${error.message || 'Unknown error'}`);
      }
    },
  });
}
