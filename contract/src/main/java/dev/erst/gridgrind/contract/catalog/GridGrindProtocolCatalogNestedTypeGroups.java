package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionInput;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorReport;
import dev.erst.gridgrind.contract.dto.FontHeightInput;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.PaneInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.PrintAreaInput;
import dev.erst.gridgrind.contract.dto.PrintScalingInput;
import dev.erst.gridgrind.contract.dto.PrintTitleColumnsInput;
import dev.erst.gridgrind.contract.dto.PrintTitleRowsInput;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.TableStyleReport;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SelectorJsonSupport;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.List;

/** Owns the nested field-shape descriptor registry used by the protocol catalog. */
final class GridGrindProtocolCatalogNestedTypeGroups {
  private GridGrindProtocolCatalogNestedTypeGroups() {}

  static final List<CatalogNestedTypeDescriptor> NESTED_TYPE_GROUPS =
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
              GridGrindResponse.NamedRangeReport.class,
              List.of(
                  descriptor(
                      GridGrindResponse.NamedRangeReport.RangeReport.class,
                      "RANGE",
                      "Exact named-range report that resolves to one typed workbook target."),
                  descriptor(
                      GridGrindResponse.NamedRangeReport.FormulaReport.class,
                      "FORMULA",
                      "Exact named-range report that remains formula-backed."))),
          nestedTypeGroup(
              "sheetProtectionReportTypes",
              GridGrindResponse.SheetProtectionReport.class,
              List.of(
                  descriptor(
                      GridGrindResponse.SheetProtectionReport.Unprotected.class,
                      "UNPROTECTED",
                      "Expect the sheet to have no protection."),
                  descriptor(
                      GridGrindResponse.SheetProtectionReport.Protected.class,
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
                      "Drawing anchor with absolute EMU coordinates and size."))),
          nestedTypeGroup(
              "cellInputTypes",
              CellInput.class,
              List.of(
                  descriptor(CellInput.Blank.class, "BLANK", "Write an empty cell."),
                  descriptor(
                      CellInput.Text.class,
                      "TEXT",
                      "Write a string cell value. Blank text is rejected; use BLANK for empty"
                          + " cells."),
                  descriptor(
                      CellInput.RichText.class,
                      "RICH_TEXT",
                      "Write a structured string cell value with an ordered rich-text run list."
                          + " Run text concatenates to the stored plain string value and each run"
                          + " may override font attributes independently."),
                  descriptor(CellInput.Numeric.class, "NUMBER", "Write a numeric cell value."),
                  descriptor(
                      CellInput.BooleanValue.class, "BOOLEAN", "Write a boolean cell value."),
                  descriptor(
                      CellInput.Date.class,
                      "DATE",
                      "Write an ISO-8601 date value."
                          + " Stored as an Excel serial number; GET_CELLS returns declaredType=NUMBER with a formatted displayValue."),
                  descriptor(
                      CellInput.DateTime.class,
                      "DATE_TIME",
                      "Write an ISO-8601 date-time value."
                          + " Stored as an Excel serial number; GET_CELLS returns declaredType=NUMBER with a formatted displayValue."),
                  descriptor(
                      CellInput.Formula.class,
                      "FORMULA",
                      "Write an Excel formula. A leading = sign is accepted and stripped automatically; the engine stores the formula without it."))),
          nestedTypeGroup(
              "hyperlinkTargetTypes",
              HyperlinkTarget.class,
              List.of(
                  descriptor(HyperlinkTarget.Url.class, "URL", "Attach an absolute URL target."),
                  descriptor(
                      HyperlinkTarget.Email.class,
                      "EMAIL",
                      "Attach an email target without the mailto: prefix."),
                  descriptor(
                      HyperlinkTarget.File.class,
                      "FILE",
                      "Attach a local or shared file path."
                          + " Accepts plain paths or file: URIs and normalizes to a plain path"
                          + " string."),
                  descriptor(
                      HyperlinkTarget.Document.class,
                      "DOCUMENT",
                      "Attach an internal workbook target."))),
          nestedTypeGroup(
              "paneTypes",
              PaneInput.class,
              List.of(
                  descriptor(
                      PaneInput.None.class, "NONE", "Sheet has no active pane split or freeze."),
                  descriptor(
                      PaneInput.Frozen.class,
                      "FROZEN",
                      "Freeze the sheet at the provided split and visible-origin coordinates."),
                  descriptor(
                      PaneInput.Split.class,
                      "SPLIT",
                      "Apply split panes with explicit split offsets, visible origin, and active pane."))),
          nestedTypeGroup(
              "sheetCopyPositionTypes",
              SheetCopyPosition.class,
              List.of(
                  descriptor(
                      SheetCopyPosition.AppendAtEnd.class,
                      "APPEND_AT_END",
                      "Place the copied sheet after every existing sheet."),
                  descriptor(
                      SheetCopyPosition.AtIndex.class,
                      "AT_INDEX",
                      "Place the copied sheet at the requested zero-based workbook position."))),
          nestedTypeGroup(
              "workbookSelectorTypes",
              dev.erst.gridgrind.contract.selector.WorkbookSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.WorkbookSelector.Current.class,
                      "Target the workbook currently being executed."))),
          nestedTypeGroup(
              "sheetSelectorTypes",
              dev.erst.gridgrind.contract.selector.SheetSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.SheetSelector.All.class,
                      "Select every sheet in workbook order."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.SheetSelector.ByName.class,
                      "Select one exact sheet by name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.SheetSelector.ByNames.class,
                      "Select one or more exact sheets by ordered name list."))),
          nestedTypeGroup(
              "cellSelectorTypes",
              dev.erst.gridgrind.contract.selector.CellSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.CellSelector.AllUsedInSheet.class,
                      "Select every physically present cell on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.CellSelector.ByAddress.class,
                      "Select one exact cell on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.CellSelector.ByAddresses.class,
                      "Select one or more exact cells on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.CellSelector.ByQualifiedAddresses.class,
                      "Select one or more exact cells across one or more sheets."))),
          nestedTypeGroup(
              "rangeSelectorTypes",
              dev.erst.gridgrind.contract.selector.RangeSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.RangeSelector.AllOnSheet.class,
                      "Select every matching range-backed structure on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.RangeSelector.ByRange.class,
                      "Select one exact rectangular range on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.RangeSelector.ByRanges.class,
                      "Select one or more exact rectangular ranges on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.RangeSelector.RectangularWindow.class,
                      "Select one rectangular window anchored at one top-left cell."))),
          nestedTypeGroup(
              "rowBandSelectorTypes",
              dev.erst.gridgrind.contract.selector.RowBandSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.RowBandSelector.Span.class,
                      "Select one inclusive zero-based row span on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.RowBandSelector.Insertion.class,
                      "Select one row insertion point plus row count on one sheet."))),
          nestedTypeGroup(
              "columnBandSelectorTypes",
              dev.erst.gridgrind.contract.selector.ColumnBandSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.ColumnBandSelector.Span.class,
                      "Select one inclusive zero-based column span on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.ColumnBandSelector.Insertion.class,
                      "Select one column insertion point plus column count on one sheet."))),
          nestedTypeGroup(
              "tableSelectorTypes",
              dev.erst.gridgrind.contract.selector.TableSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.TableSelector.All.class,
                      "Select every table in workbook order."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.TableSelector.ByName.class,
                      "Select one workbook-global table by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.TableSelector.ByNames.class,
                      "Select one or more workbook-global tables by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.TableSelector.ByNameOnSheet.class,
                      "Select one workbook-global table by exact name and expected owning sheet."))),
          nestedTypeGroup(
              "tableRowSelectorTypes",
              dev.erst.gridgrind.contract.selector.TableRowSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.TableRowSelector.AllRows.class,
                      "Select every logical data row in one selected table."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.TableRowSelector.ByIndex.class,
                      "Select one zero-based data row by index in one selected table."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.TableRowSelector.ByKeyCell.class,
                      "Select one logical data row by matching one key-column cell value."))),
          nestedTypeGroup(
              "tableCellSelectorTypes",
              dev.erst.gridgrind.contract.selector.TableCellSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.TableCellSelector.ByColumnName.class,
                      "Select one logical cell within one selected table row by column name."))),
          nestedTypeGroup(
              "namedRangeRefSelectorTypes",
              dev.erst.gridgrind.contract.selector.NamedRangeSelector.Ref.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByName.class,
                      "Match a named range reference across all scopes by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope.class,
                      "Match one workbook-scoped named range reference by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope.class,
                      "Match one sheet-scoped named range reference on one sheet."))),
          nestedTypeGroup(
              "namedRangeScopeTypes",
              NamedRangeScope.class,
              List.of(
                  descriptor(NamedRangeScope.Workbook.class, "WORKBOOK", "Target workbook scope."),
                  descriptor(
                      NamedRangeScope.Sheet.class, "SHEET", "Target one specific sheet scope."))),
          nestedTypeGroup(
              "namedRangeSelectorTypes",
              dev.erst.gridgrind.contract.selector.NamedRangeSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.All.class,
                      "Select every user-facing named range."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf.class,
                      "Select the union of one or more explicit named-range references."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByName.class,
                      "Match a named range across all scopes by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByNames.class,
                      "Match named ranges across all scopes by exact name set."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope.class,
                      "Match the workbook-scoped named range with the exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope.class,
                      "Match the sheet-scoped named range on one sheet."))),
          nestedTypeGroup(
              "drawingObjectSelectorTypes",
              dev.erst.gridgrind.contract.selector.DrawingObjectSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.DrawingObjectSelector.AllOnSheet.class,
                      "Select every drawing object on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.DrawingObjectSelector.ByName.class,
                      "Select one drawing object by exact sheet-local object name."))),
          nestedTypeGroup(
              "chartSelectorTypes",
              dev.erst.gridgrind.contract.selector.ChartSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.ChartSelector.AllOnSheet.class,
                      "Select every chart on one sheet."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.ChartSelector.ByName.class,
                      "Select one chart by exact sheet-local chart name."))),
          nestedTypeGroup(
              "pivotTableSelectorTypes",
              dev.erst.gridgrind.contract.selector.PivotTableSelector.class,
              List.of(
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.PivotTableSelector.All.class,
                      "Select every pivot table in workbook order."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.PivotTableSelector.ByName.class,
                      "Select one workbook-global pivot table by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.PivotTableSelector.ByNames.class,
                      "Select one or more workbook-global pivot tables by exact name."),
                  selectorDescriptor(
                      dev.erst.gridgrind.contract.selector.PivotTableSelector.ByNameOnSheet.class,
                      "Select one workbook-global pivot table by exact name and expected owning sheet."))),
          nestedTypeGroup(
              "drawingAnchorInputTypes",
              DrawingAnchorInput.class,
              List.of(
                  descriptor(
                      DrawingAnchorInput.TwoCell.class,
                      "TWO_CELL",
                      "Authored two-cell drawing anchor with explicit start and end markers."
                          + " behavior defaults to MOVE_AND_RESIZE when omitted.",
                      "behavior"))),
          nestedTypeGroup(
              "chartPlotInputTypes",
              ChartInput.Plot.class,
              List.of(
                  descriptor(
                      ChartInput.Area.class,
                      "AREA",
                      "Authored area-chart plot."
                          + " varyColors, grouping, and axes default when omitted.",
                      "varyColors",
                      "grouping",
                      "axes"),
                  descriptor(
                      ChartInput.Area3D.class,
                      "AREA_3D",
                      "Authored 3D area-chart plot."
                          + " varyColors, grouping, gapDepth, and axes default when omitted.",
                      "varyColors",
                      "grouping",
                      "gapDepth",
                      "axes"),
                  descriptor(
                      ChartInput.Bar.class,
                      "BAR",
                      "Authored bar-chart plot."
                          + " varyColors, barDirection, grouping, gapWidth, overlap,"
                          + " and axes default when omitted.",
                      "varyColors",
                      "barDirection",
                      "grouping",
                      "gapWidth",
                      "overlap",
                      "axes"),
                  descriptor(
                      ChartInput.Bar3D.class,
                      "BAR_3D",
                      "Authored 3D bar-chart plot."
                          + " varyColors, barDirection, grouping, gapDepth, gapWidth,"
                          + " shape, and axes default when omitted.",
                      "varyColors",
                      "barDirection",
                      "grouping",
                      "gapDepth",
                      "gapWidth",
                      "shape",
                      "axes"),
                  descriptor(
                      ChartInput.Doughnut.class,
                      "DOUGHNUT",
                      "Authored doughnut-chart plot."
                          + " varyColors, firstSliceAngle, and holeSize are optional.",
                      "varyColors",
                      "firstSliceAngle",
                      "holeSize"),
                  descriptor(
                      ChartInput.Line.class,
                      "LINE",
                      "Authored line-chart plot."
                          + " varyColors, grouping, and axes default when omitted.",
                      "varyColors",
                      "grouping",
                      "axes"),
                  descriptor(
                      ChartInput.Line3D.class,
                      "LINE_3D",
                      "Authored 3D line-chart plot."
                          + " varyColors, grouping, gapDepth, and axes default when omitted.",
                      "varyColors",
                      "grouping",
                      "gapDepth",
                      "axes"),
                  descriptor(
                      ChartInput.Pie.class,
                      "PIE",
                      "Authored pie-chart plot." + " varyColors and firstSliceAngle are optional.",
                      "varyColors",
                      "firstSliceAngle"),
                  descriptor(
                      ChartInput.Pie3D.class,
                      "PIE_3D",
                      "Authored 3D pie-chart plot." + " varyColors defaults when omitted.",
                      "varyColors"),
                  descriptor(
                      ChartInput.Radar.class,
                      "RADAR",
                      "Authored radar-chart plot."
                          + " varyColors, style, and axes default when omitted.",
                      "varyColors",
                      "style",
                      "axes"),
                  descriptor(
                      ChartInput.Scatter.class,
                      "SCATTER",
                      "Authored scatter-chart plot."
                          + " varyColors, style, and axes default when omitted.",
                      "varyColors",
                      "style",
                      "axes"),
                  descriptor(
                      ChartInput.Surface.class,
                      "SURFACE",
                      "Authored surface-chart plot."
                          + " varyColors, wireframe, and axes default when omitted.",
                      "varyColors",
                      "wireframe",
                      "axes"),
                  descriptor(
                      ChartInput.Surface3D.class,
                      "SURFACE_3D",
                      "Authored 3D surface-chart plot."
                          + " varyColors, wireframe, and axes default when omitted.",
                      "varyColors",
                      "wireframe",
                      "axes"))),
          nestedTypeGroup(
              "chartTitleInputTypes",
              ChartInput.Title.class,
              List.of(
                  descriptor(
                      ChartInput.Title.None.class, "NONE", "Remove any chart or series title."),
                  descriptor(ChartInput.Title.Text.class, "TEXT", "Use one explicit static title."),
                  descriptor(
                      ChartInput.Title.Formula.class,
                      "FORMULA",
                      "Bind the chart or series title to one workbook formula that resolves"
                          + " to one cell."))),
          nestedTypeGroup(
              "chartLegendInputTypes",
              ChartInput.Legend.class,
              List.of(
                  descriptor(ChartInput.Legend.Hidden.class, "HIDDEN", "Hide the legend entirely."),
                  descriptor(
                      ChartInput.Legend.Visible.class,
                      "VISIBLE",
                      "Show the legend at one explicit position."))),
          nestedTypeGroup(
              "chartDataSourceInputTypes",
              ChartInput.DataSource.class,
              List.of(
                  descriptor(
                      ChartInput.DataSource.Reference.class,
                      "REFERENCE",
                      "Workbook-backed chart source formula or defined name."),
                  descriptor(
                      ChartInput.DataSource.StringLiteral.class,
                      "STRING_LITERAL",
                      "Literal string chart source stored directly in the chart part."),
                  descriptor(
                      ChartInput.DataSource.NumericLiteral.class,
                      "NUMERIC_LITERAL",
                      "Literal numeric chart source stored directly in the chart part."))),
          nestedTypeGroup(
              "pivotTableSourceTypes",
              PivotTableInput.Source.class,
              List.of(
                  descriptor(
                      PivotTableInput.Source.Range.class,
                      "RANGE",
                      "Use one explicit contiguous sheet range with the header row in the first"
                          + " row."),
                  descriptor(
                      PivotTableInput.Source.NamedRange.class,
                      "NAMED_RANGE",
                      "Use one existing workbook- or sheet-scoped named range as the pivot"
                          + " source."),
                  descriptor(
                      PivotTableInput.Source.Table.class,
                      "TABLE",
                      "Use one existing workbook-global table as the pivot source."))),
          nestedTypeGroup(
              "fontHeightTypes",
              FontHeightInput.class,
              List.of(
                  descriptor(
                      FontHeightInput.Points.class,
                      "POINTS",
                      "Specify font height in point units."
                          + " Write format: {\"type\":\"POINTS\",\"points\":13}."
                          + " Read-back (GET_CELLS, GET_WINDOW): style.fontHeight is"
                          + " {\"twips\":260,\"points\":13} with both fields present,"
                          + " not this discriminated type format."),
                  descriptor(
                      FontHeightInput.Twips.class,
                      "TWIPS",
                      "Specify font height in exact twips (20 twips = 1 point)."
                          + " Write format: {\"type\":\"TWIPS\",\"twips\":260}."
                          + " Read-back returns the same plain object shape as POINTS."))),
          nestedTypeGroup(
              "dataValidationRuleTypes",
              DataValidationRuleInput.class,
              List.of(
                  descriptor(
                      DataValidationRuleInput.ExplicitList.class,
                      "EXPLICIT_LIST",
                      "Allow only one of the supplied explicit values."
                          + " An empty values array preserves Excel's explicit-empty-list state."),
                  descriptor(
                      DataValidationRuleInput.FormulaList.class,
                      "FORMULA_LIST",
                      "Allow values from a formula-driven list expression."),
                  descriptor(
                      DataValidationRuleInput.WholeNumber.class,
                      "WHOLE_NUMBER",
                      "Apply a whole-number comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.DecimalNumber.class,
                      "DECIMAL_NUMBER",
                      "Apply a decimal-number comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.DateRule.class,
                      "DATE",
                      "Apply a date comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.TimeRule.class,
                      "TIME",
                      "Apply a time comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.TextLength.class,
                      "TEXT_LENGTH",
                      "Apply a text-length comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.CustomFormula.class,
                      "CUSTOM_FORMULA",
                      "Allow values that satisfy a custom formula."))),
          nestedTypeGroup(
              "autofilterFilterCriterionTypes",
              AutofilterFilterCriterionInput.class,
              List.of(
                  descriptor(
                      AutofilterFilterCriterionInput.Values.class,
                      "VALUES",
                      "Retain rows whose cell values match one or more explicit values."),
                  descriptor(
                      AutofilterFilterCriterionInput.Custom.class,
                      "CUSTOM",
                      "Retain rows that satisfy one or two comparator-based custom conditions."),
                  descriptor(
                      AutofilterFilterCriterionInput.Dynamic.class,
                      "DYNAMIC",
                      "Retain rows using one dynamic-date or moving-window autofilter rule.",
                      "value",
                      "maxValue"),
                  descriptor(
                      AutofilterFilterCriterionInput.Top10.class,
                      "TOP10",
                      "Retain top or bottom N or percent values."),
                  descriptor(
                      AutofilterFilterCriterionInput.Color.class,
                      "COLOR",
                      "Retain rows matching one cell color or font color criterion."),
                  descriptor(
                      AutofilterFilterCriterionInput.Icon.class,
                      "ICON",
                      "Retain rows matching one icon-set member."))),
          nestedTypeGroup(
              "conditionalFormattingRuleTypes",
              ConditionalFormattingRuleInput.class,
              List.of(
                  descriptor(
                      ConditionalFormattingRuleInput.FormulaRule.class,
                      "FORMULA_RULE",
                      "Apply one formula-driven conditional-formatting rule."
                          + " Supply one differential style."),
                  descriptor(
                      ConditionalFormattingRuleInput.CellValueRule.class,
                      "CELL_VALUE_RULE",
                      "Apply one cell-value comparison conditional-formatting rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN."
                          + " Supply one differential style.",
                      "formula2"),
                  descriptor(
                      ConditionalFormattingRuleInput.ColorScaleRule.class,
                      "COLOR_SCALE_RULE",
                      "Apply one color-scale conditional-formatting rule with ordered thresholds"
                          + " and colors."),
                  descriptor(
                      ConditionalFormattingRuleInput.DataBarRule.class,
                      "DATA_BAR_RULE",
                      "Apply one data-bar conditional-formatting rule with explicit thresholds"
                          + " and widths."),
                  descriptor(
                      ConditionalFormattingRuleInput.IconSetRule.class,
                      "ICON_SET_RULE",
                      "Apply one icon-set conditional-formatting rule with authored thresholds."),
                  descriptor(
                      ConditionalFormattingRuleInput.Top10Rule.class,
                      "TOP10_RULE",
                      "Apply one top/bottom-N conditional-formatting rule with a differential"
                          + " style."))),
          nestedTypeGroup(
              "printAreaTypes",
              PrintAreaInput.class,
              List.of(
                  descriptor(
                      PrintAreaInput.None.class, "NONE", "Sheet has no explicit print area."),
                  descriptor(
                      PrintAreaInput.Range.class,
                      "RANGE",
                      "Sheet prints the provided rectangular A1-style range."))),
          nestedTypeGroup(
              "printScalingTypes",
              PrintScalingInput.class,
              List.of(
                  descriptor(
                      PrintScalingInput.Automatic.class,
                      "AUTOMATIC",
                      "Sheet uses Excel's default scaling instead of fit-to-page counts."),
                  descriptor(
                      PrintScalingInput.Fit.class,
                      "FIT",
                      "Sheet fits printed content into the provided page counts."
                          + " A value of 0 on one axis keeps that axis unconstrained."))),
          nestedTypeGroup(
              "printTitleRowsTypes",
              PrintTitleRowsInput.class,
              List.of(
                  descriptor(
                      PrintTitleRowsInput.None.class,
                      "NONE",
                      "Sheet has no repeating print-title rows."),
                  descriptor(
                      PrintTitleRowsInput.Band.class,
                      "BAND",
                      "Sheet repeats the provided inclusive zero-based row band on every printed page."))),
          nestedTypeGroup(
              "printTitleColumnsTypes",
              PrintTitleColumnsInput.class,
              List.of(
                  descriptor(
                      PrintTitleColumnsInput.None.class,
                      "NONE",
                      "Sheet has no repeating print-title columns."),
                  descriptor(
                      PrintTitleColumnsInput.Band.class,
                      "BAND",
                      "Sheet repeats the provided inclusive zero-based column band on every printed page."))),
          nestedTypeGroup(
              "tableStyleTypes",
              TableStyleInput.class,
              List.of(
                  descriptor(TableStyleInput.None.class, "NONE", "Clear table style metadata."),
                  descriptor(
                      TableStyleInput.Named.class,
                      "NAMED",
                      "Apply one named workbook table style with explicit stripe and emphasis"
                          + " flags."))),
          nestedTypeGroup(
              "calculationStrategyTypes",
              CalculationStrategyInput.class,
              List.of(
                  descriptor(
                      CalculationStrategyInput.DoNotCalculate.class,
                      "DO_NOT_CALCULATE",
                      "Skip immediate server-side formula calculation."),
                  descriptor(
                      CalculationStrategyInput.EvaluateAll.class,
                      "EVALUATE_ALL",
                      "Preflight and evaluate every formula cell after mutation steps complete."),
                  descriptor(
                      CalculationStrategyInput.EvaluateTargets.class,
                      "EVALUATE_TARGETS",
                      "Preflight and evaluate the explicit formula-cell target list only.",
                      "cells"),
                  descriptor(
                      CalculationStrategyInput.ClearCachesOnly.class,
                      "CLEAR_CACHES_ONLY",
                      "Strip persisted formula caches without running immediate evaluation."))));

  private static CatalogNestedTypeDescriptor nestedTypeGroup(
      String group, Class<?> sealedType, List<CatalogTypeDescriptor> typeDescriptors) {
    return GridGrindProtocolCatalog.nestedTypeGroup(group, sealedType, typeDescriptors);
  }

  private static CatalogTypeDescriptor descriptor(
      Class<? extends Record> recordType, String id, String summary, String... optionalFields) {
    return GridGrindProtocolCatalog.descriptor(recordType, id, summary, optionalFields);
  }

  private static <T extends Record & Selector> CatalogTypeDescriptor selectorDescriptor(
      Class<T> recordType, String summary, String... optionalFields) {
    return descriptor(
        recordType, SelectorJsonSupport.typeIdFor(recordType), summary, optionalFields);
  }
}
