package dev.erst.gridgrind.contract.catalog;

import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.descriptor;
import static dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalogNestedTypeGroupSupport.nestedTypeGroup;

import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.DrawingAnchorReport;
import dev.erst.gridgrind.contract.dto.GridGrindWorkbookSurfaceReports;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.TableStyleReport;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.List;

/** Owns one focused subset of nested-type group descriptors for the protocol catalog. */
final class GridGrindProtocolCatalogSourceAndReportNestedTypeGroups {
  private GridGrindProtocolCatalogSourceAndReportNestedTypeGroups() {}

  static final List<CatalogNestedTypeDescriptor> SOURCE_AND_REPORT_GROUPS =
      List.of(
          nestedTypeGroup(
              "expectedCellValueTypes",
              dev.erst.gridgrind.contract.assertion.ExpectedCellValue.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.assertion.ExpectedCellValue.Blank.class,
                      "BLANK",
                      "Require the effective cell value to be blank."),
                  descriptor(
                      dev.erst.gridgrind.contract.assertion.ExpectedCellValue.Text.class,
                      "TEXT",
                      "Require the effective cell value to be one exact string."),
                  descriptor(
                      dev.erst.gridgrind.contract.assertion.ExpectedCellValue.NumericValue.class,
                      "NUMBER",
                      "Require the effective cell value to be one exact finite number."),
                  descriptor(
                      dev.erst.gridgrind.contract.assertion.ExpectedCellValue.BooleanValue.class,
                      "BOOLEAN",
                      "Require the effective cell value to be true or false."),
                  descriptor(
                      dev.erst.gridgrind.contract.assertion.ExpectedCellValue.ErrorValue.class,
                      "ERROR",
                      "Require the effective cell value to be one exact Excel error string."))),
          nestedTypeGroup(
              "textSourceTypes",
              TextSourceInput.class,
              List.of(
                  descriptor(
                      TextSourceInput.Inline.class,
                      "INLINE",
                      "Embed UTF-8 text directly in the request JSON."),
                  descriptor(
                      TextSourceInput.Utf8File.class,
                      "UTF8_FILE",
                      "Load UTF-8 text from one file path."
                          + " "
                          + GridGrindContractText.requestOwnedPathResolutionSummary()),
                  descriptor(
                      TextSourceInput.StandardInput.class,
                      "STANDARD_INPUT",
                      "Load UTF-8 text from the execution transport's bound standard input"
                          + " bytes."))),
          nestedTypeGroup(
              "binarySourceTypes",
              BinarySourceInput.class,
              List.of(
                  descriptor(
                      BinarySourceInput.InlineBase64.class,
                      "INLINE_BASE64",
                      "Embed base64-encoded binary content directly in the request JSON."),
                  descriptor(
                      BinarySourceInput.File.class,
                      "FILE",
                      "Load binary content from one file path."
                          + " "
                          + GridGrindContractText.requestOwnedPathResolutionSummary()),
                  descriptor(
                      BinarySourceInput.StandardInput.class,
                      "STANDARD_INPUT",
                      "Load binary content from the execution transport's bound standard input"
                          + " bytes."))),
          nestedTypeGroup(
              "namedRangeReportTypes",
              GridGrindWorkbookSurfaceReports.NamedRangeReport.class,
              List.of(
                  descriptor(
                      GridGrindWorkbookSurfaceReports.NamedRangeReport.RangeReport.class,
                      "RANGE",
                      "Exact named-range report that resolves to one typed workbook target."),
                  descriptor(
                      GridGrindWorkbookSurfaceReports.NamedRangeReport.FormulaReport.class,
                      "FORMULA",
                      "Exact named-range report that remains formula-backed."))),
          nestedTypeGroup(
              "sheetProtectionReportTypes",
              GridGrindWorkbookSurfaceReports.SheetProtectionReport.class,
              List.of(
                  descriptor(
                      GridGrindWorkbookSurfaceReports.SheetProtectionReport.Unprotected.class,
                      "UNPROTECTED",
                      "Expect the sheet to have no protection."),
                  descriptor(
                      GridGrindWorkbookSurfaceReports.SheetProtectionReport.Protected.class,
                      "PROTECTED",
                      "Expect the sheet to be protected with explicit lock settings."))),
          nestedTypeGroup(
              "tableStyleReportTypes",
              TableStyleReport.class,
              List.of(
                  descriptor(
                      TableStyleReport.None.class,
                      "NONE",
                      "Expect the table to carry no persisted style."),
                  descriptor(
                      TableStyleReport.Named.class,
                      "NAMED",
                      "Expect the table to carry one named style plus stripe/emphasis flags."))),
          nestedTypeGroup(
              "pivotTableReportTypes",
              PivotTableReport.class,
              List.of(
                  descriptor(
                      PivotTableReport.Supported.class,
                      "SUPPORTED",
                      "Exact supported pivot-table report."),
                  descriptor(
                      PivotTableReport.Unsupported.class,
                      "UNSUPPORTED",
                      "Exact unsupported pivot-table report preserved from the workbook."))),
          nestedTypeGroup(
              "pivotTableReportSourceTypes",
              PivotTableReport.Source.class,
              List.of(
                  descriptor(
                      PivotTableReport.Source.Range.class,
                      "RANGE",
                      "Pivot source resolved from one sheet range."),
                  descriptor(
                      PivotTableReport.Source.NamedRange.class,
                      "NAMED_RANGE",
                      "Pivot source resolved from one named range."),
                  descriptor(
                      PivotTableReport.Source.Table.class,
                      "TABLE",
                      "Pivot source resolved from one workbook table."))),
          nestedTypeGroup(
              "chartPlotReportTypes",
              ChartReport.Plot.class,
              List.of(
                  descriptor(ChartReport.Area.class, "AREA", "Exact area-chart plot report."),
                  descriptor(
                      ChartReport.Area3D.class,
                      "AREA_3D",
                      "Exact 3D area-chart plot report.",
                      "gapDepth"),
                  descriptor(
                      ChartReport.Bar.class,
                      "BAR",
                      "Exact bar-chart plot report.",
                      "gapWidth",
                      "overlap"),
                  descriptor(
                      ChartReport.Bar3D.class,
                      "BAR_3D",
                      "Exact 3D bar-chart plot report.",
                      "gapDepth",
                      "gapWidth",
                      "shape"),
                  descriptor(
                      ChartReport.Doughnut.class,
                      "DOUGHNUT",
                      "Exact doughnut-chart plot report.",
                      "firstSliceAngle",
                      "holeSize"),
                  descriptor(ChartReport.Line.class, "LINE", "Exact line-chart plot report."),
                  descriptor(
                      ChartReport.Line3D.class,
                      "LINE_3D",
                      "Exact 3D line-chart plot report.",
                      "gapDepth"),
                  descriptor(
                      ChartReport.Pie.class,
                      "PIE",
                      "Exact pie-chart plot report.",
                      "firstSliceAngle"),
                  descriptor(ChartReport.Pie3D.class, "PIE_3D", "Exact 3D pie-chart plot report."),
                  descriptor(ChartReport.Radar.class, "RADAR", "Exact radar-chart plot report."),
                  descriptor(
                      ChartReport.Scatter.class, "SCATTER", "Exact scatter-chart plot report."),
                  descriptor(
                      ChartReport.Surface.class, "SURFACE", "Exact surface-chart plot report."),
                  descriptor(
                      ChartReport.Surface3D.class,
                      "SURFACE_3D",
                      "Exact 3D surface-chart plot report."),
                  descriptor(
                      ChartReport.Unsupported.class,
                      "UNSUPPORTED",
                      "Exact unsupported chart plot report preserved from the workbook."))),
          nestedTypeGroup(
              "chartTitleReportTypes",
              ChartReport.Title.class,
              List.of(
                  descriptor(ChartReport.Title.None.class, "NONE", "No title is present."),
                  descriptor(ChartReport.Title.Text.class, "TEXT", "Static title text."),
                  descriptor(
                      ChartReport.Title.Formula.class,
                      "FORMULA",
                      "Formula-backed title with cached text."))),
          nestedTypeGroup(
              "chartLegendReportTypes",
              ChartReport.Legend.class,
              List.of(
                  descriptor(ChartReport.Legend.Hidden.class, "HIDDEN", "No legend is present."),
                  descriptor(
                      ChartReport.Legend.Visible.class,
                      "VISIBLE",
                      "Visible legend at one persisted position."))),
          nestedTypeGroup(
              "chartDataSourceReportTypes",
              ChartReport.DataSource.class,
              List.of(
                  descriptor(
                      ChartReport.DataSource.StringReference.class,
                      "STRING_REFERENCE",
                      "Formula-backed string chart source plus cached values."),
                  descriptor(
                      ChartReport.DataSource.NumericReference.class,
                      "NUMERIC_REFERENCE",
                      "Formula-backed numeric chart source plus cached values."),
                  descriptor(
                      ChartReport.DataSource.StringLiteral.class,
                      "STRING_LITERAL",
                      "Literal string chart source stored directly in the chart part."),
                  descriptor(
                      ChartReport.DataSource.NumericLiteral.class,
                      "NUMERIC_LITERAL",
                      "Literal numeric chart source stored directly in the chart part."))),
          nestedTypeGroup(
              "drawingAnchorReportTypes",
              DrawingAnchorReport.class,
              List.of(
                  descriptor(
                      DrawingAnchorReport.TwoCell.class,
                      "TWO_CELL",
                      "Drawing anchor spanning one start and end marker."),
                  descriptor(
                      DrawingAnchorReport.OneCell.class,
                      "ONE_CELL",
                      "Drawing anchor with one start marker plus explicit size."),
                  descriptor(
                      DrawingAnchorReport.Absolute.class,
                      "ABSOLUTE",
                      "Drawing anchor with absolute EMU coordinates and size."))));
}
