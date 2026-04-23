package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.dto.ArrayFormulaInput;
import dev.erst.gridgrind.contract.dto.ArrayFormulaReport;
import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnInput;
import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortConditionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateInput;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.CellAlignmentInput;
import dev.erst.gridgrind.contract.dto.CellAlignmentReport;
import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideInput;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellColorReport;
import dev.erst.gridgrind.contract.dto.CellFillInput;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellGradientFillInput;
import dev.erst.gridgrind.contract.dto.CellGradientFillReport;
import dev.erst.gridgrind.contract.dto.CellGradientStopInput;
import dev.erst.gridgrind.contract.dto.CellGradientStopReport;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellProtectionInput;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.CellStyleInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.ColorInput;
import dev.erst.gridgrind.contract.dto.CommentAnchorInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdInput;
import dev.erst.gridgrind.contract.dto.CustomXmlDataBindingReport;
import dev.erst.gridgrind.contract.dto.CustomXmlExportReport;
import dev.erst.gridgrind.contract.dto.CustomXmlImportInput;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingReport;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderSideInput;
import dev.erst.gridgrind.contract.dto.DifferentialStyleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorReport;
import dev.erst.gridgrind.contract.dto.DrawingMarkerInput;
import dev.erst.gridgrind.contract.dto.DrawingMarkerReport;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.ExecutionJournalInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FontHeightInput;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.FormulaExternalWorkbookInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfFunctionInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfToolpackInput;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextInput;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.IgnoredErrorInput;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionReport;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureInput;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureReport;
import dev.erst.gridgrind.contract.dto.PaneInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.PrintAreaInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.PrintMarginsInput;
import dev.erst.gridgrind.contract.dto.PrintScalingInput;
import dev.erst.gridgrind.contract.dto.PrintSetupInput;
import dev.erst.gridgrind.contract.dto.PrintTitleColumnsInput;
import dev.erst.gridgrind.contract.dto.PrintTitleRowsInput;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.contract.dto.SheetDefaultsInput;
import dev.erst.gridgrind.contract.dto.SheetDisplayInput;
import dev.erst.gridgrind.contract.dto.SheetOutlineSummaryInput;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import dev.erst.gridgrind.contract.dto.TableColumnInput;
import dev.erst.gridgrind.contract.dto.TableColumnReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.TableStyleReport;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.List;

/** Owns the nested/plain field-shape descriptor registry used by the protocol catalog. */
@SuppressWarnings("PMD.ExcessiveImports")
final class GridGrindProtocolCatalogFieldGroupSupport {
  private GridGrindProtocolCatalogFieldGroupSupport() {}

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
                  descriptor(
                      dev.erst.gridgrind.contract.selector.WorkbookSelector.Current.class,
                      "CURRENT",
                      "Target the workbook currently being executed."))),
          nestedTypeGroup(
              "sheetSelectorTypes",
              dev.erst.gridgrind.contract.selector.SheetSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.SheetSelector.All.class,
                      "ALL",
                      "Select every sheet in workbook order."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.SheetSelector.ByName.class,
                      "BY_NAME",
                      "Select one exact sheet by name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.SheetSelector.ByNames.class,
                      "BY_NAMES",
                      "Select one or more exact sheets by ordered name list."))),
          nestedTypeGroup(
              "cellSelectorTypes",
              dev.erst.gridgrind.contract.selector.CellSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.CellSelector.AllUsedInSheet.class,
                      "ALL_USED_IN_SHEET",
                      "Select every physically present cell on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.CellSelector.ByAddress.class,
                      "BY_ADDRESS",
                      "Select one exact cell on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.CellSelector.ByAddresses.class,
                      "BY_ADDRESSES",
                      "Select one or more exact cells on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.CellSelector.ByQualifiedAddresses.class,
                      "BY_QUALIFIED_ADDRESSES",
                      "Select one or more exact cells across one or more sheets."))),
          nestedTypeGroup(
              "rangeSelectorTypes",
              dev.erst.gridgrind.contract.selector.RangeSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.RangeSelector.AllOnSheet.class,
                      "ALL_ON_SHEET",
                      "Select every matching range-backed structure on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.RangeSelector.ByRange.class,
                      "BY_RANGE",
                      "Select one exact rectangular range on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.RangeSelector.ByRanges.class,
                      "BY_RANGES",
                      "Select one or more exact rectangular ranges on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.RangeSelector.RectangularWindow.class,
                      "RECTANGULAR_WINDOW",
                      "Select one rectangular window anchored at one top-left cell."))),
          nestedTypeGroup(
              "rowBandSelectorTypes",
              dev.erst.gridgrind.contract.selector.RowBandSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.RowBandSelector.Span.class,
                      "SPAN",
                      "Select one inclusive zero-based row span on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.RowBandSelector.Insertion.class,
                      "INSERTION",
                      "Select one row insertion point plus row count on one sheet."))),
          nestedTypeGroup(
              "columnBandSelectorTypes",
              dev.erst.gridgrind.contract.selector.ColumnBandSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.ColumnBandSelector.Span.class,
                      "SPAN",
                      "Select one inclusive zero-based column span on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.ColumnBandSelector.Insertion.class,
                      "INSERTION",
                      "Select one column insertion point plus column count on one sheet."))),
          nestedTypeGroup(
              "tableSelectorTypes",
              dev.erst.gridgrind.contract.selector.TableSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.TableSelector.All.class,
                      "ALL",
                      "Select every table in workbook order."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.TableSelector.ByName.class,
                      "BY_NAME",
                      "Select one workbook-global table by exact name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.TableSelector.ByNames.class,
                      "BY_NAMES",
                      "Select one or more workbook-global tables by exact name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.TableSelector.ByNameOnSheet.class,
                      "BY_NAME_ON_SHEET",
                      "Select one workbook-global table by exact name and expected owning sheet."))),
          nestedTypeGroup(
              "tableRowSelectorTypes",
              dev.erst.gridgrind.contract.selector.TableRowSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.TableRowSelector.AllRows.class,
                      "ALL_ROWS",
                      "Select every logical data row in one selected table."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.TableRowSelector.ByIndex.class,
                      "BY_INDEX",
                      "Select one zero-based data row by index in one selected table."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.TableRowSelector.ByKeyCell.class,
                      "BY_KEY_CELL",
                      "Select one logical data row by matching one key-column cell value."))),
          nestedTypeGroup(
              "tableCellSelectorTypes",
              dev.erst.gridgrind.contract.selector.TableCellSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.TableCellSelector.ByColumnName.class,
                      "BY_COLUMN_NAME",
                      "Select one logical cell within one selected table row by column name."))),
          nestedTypeGroup(
              "namedRangeRefSelectorTypes",
              dev.erst.gridgrind.contract.selector.NamedRangeSelector.Ref.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByName.class,
                      "BY_NAME",
                      "Match a named range reference across all scopes by exact name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope.class,
                      "WORKBOOK_SCOPE",
                      "Match one workbook-scoped named range reference by exact name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope.class,
                      "SHEET_SCOPE",
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
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.All.class,
                      "ALL",
                      "Select every user-facing named range."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf.class,
                      "ANY_OF",
                      "Select the union of one or more explicit named-range references."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByName.class,
                      "BY_NAME",
                      "Match a named range across all scopes by exact name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByNames.class,
                      "BY_NAMES",
                      "Match named ranges across all scopes by exact name set."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope.class,
                      "WORKBOOK_SCOPE",
                      "Match the workbook-scoped named range with the exact name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope.class,
                      "SHEET_SCOPE",
                      "Match the sheet-scoped named range on one sheet."))),
          nestedTypeGroup(
              "drawingObjectSelectorTypes",
              dev.erst.gridgrind.contract.selector.DrawingObjectSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.DrawingObjectSelector.AllOnSheet.class,
                      "ALL_ON_SHEET",
                      "Select every drawing object on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.DrawingObjectSelector.ByName.class,
                      "BY_NAME",
                      "Select one drawing object by exact sheet-local object name."))),
          nestedTypeGroup(
              "chartSelectorTypes",
              dev.erst.gridgrind.contract.selector.ChartSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.ChartSelector.AllOnSheet.class,
                      "ALL_ON_SHEET",
                      "Select every chart on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.ChartSelector.ByName.class,
                      "BY_NAME",
                      "Select one chart by exact sheet-local chart name."))),
          nestedTypeGroup(
              "pivotTableSelectorTypes",
              dev.erst.gridgrind.contract.selector.PivotTableSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.PivotTableSelector.All.class,
                      "ALL",
                      "Select every pivot table in workbook order."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.PivotTableSelector.ByName.class,
                      "BY_NAME",
                      "Select one workbook-global pivot table by exact name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.PivotTableSelector.ByNames.class,
                      "BY_NAMES",
                      "Select one or more workbook-global pivot tables by exact name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.PivotTableSelector.ByNameOnSheet.class,
                      "BY_NAME_ON_SHEET",
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
  static final List<CatalogPlainTypeDescriptor> PLAIN_TYPE_DESCRIPTORS =
      List.of(
          plainTypeDescriptor(
              "executionJournalType",
              ExecutionJournal.class,
              "ExecutionJournal",
              "Structured execution telemetry returned on every success and failure response,"
                  + " including validation, open, calculation, step, persistence, and close"
                  + " phases.",
              List.of(
                  "planId",
                  "level",
                  "source",
                  "persistence",
                  "validation",
                  "open",
                  "calculation",
                  "persistencePhase",
                  "close",
                  "steps",
                  "warnings",
                  "outcome",
                  "events")),
          plainTypeDescriptor(
              "executionJournalSourceSummaryType",
              ExecutionJournal.SourceSummary.class,
              "ExecutionJournalSourceSummary",
              "Journal summary of the authored workbook source.",
              List.of("path")),
          plainTypeDescriptor(
              "executionJournalPersistenceSummaryType",
              ExecutionJournal.PersistenceSummary.class,
              "ExecutionJournalPersistenceSummary",
              "Journal summary of the authored persistence policy.",
              List.of("path")),
          plainTypeDescriptor(
              "executionJournalPhaseType",
              ExecutionJournal.Phase.class,
              "ExecutionJournalPhase",
              "One timed execution phase with status, timestamps, and duration.",
              List.of("startedAt", "finishedAt")),
          plainTypeDescriptor(
              "executionJournalStepType",
              ExecutionJournal.Step.class,
              "ExecutionJournalStep",
              "Per-step execution telemetry with resolved targets, timing, outcome,"
                  + " and optional failure detail.",
              List.of("failure")),
          plainTypeDescriptor(
              "executionJournalTargetType",
              ExecutionJournal.Target.class,
              "ExecutionJournalTarget",
              "One canonical target label recorded inside a step journal.",
              List.of()),
          plainTypeDescriptor(
              "executionJournalFailureClassificationType",
              ExecutionJournal.FailureClassification.class,
              "ExecutionJournalFailureClassification",
              "Structured problem-code classification for one failed step.",
              List.of()),
          plainTypeDescriptor(
              "executionJournalCalculationType",
              ExecutionJournal.Calculation.class,
              "ExecutionJournalCalculation",
              "Top-level calculation preflight and execution timings for one request.",
              List.of()),
          plainTypeDescriptor(
              "executionJournalOutcomeType",
              ExecutionJournal.Outcome.class,
              "ExecutionJournalOutcome",
              "Final outcome summary for one execution journal.",
              List.of("failedStepIndex", "failedStepId", "failureCode")),
          plainTypeDescriptor(
              "executionJournalEventType",
              ExecutionJournal.Event.class,
              "ExecutionJournalEvent",
              "Fine-grained verbose execution event emitted for live CLI rendering.",
              List.of("stepIndex", "stepId")),
          plainTypeDescriptor(
              "requestWarningType",
              RequestWarning.class,
              "RequestWarning",
              "Non-fatal authored-plan warning surfaced on success and echoed inside the execution journal.",
              List.of()),
          plainTypeDescriptor(
              "executionPolicyInputType",
              ExecutionPolicyInput.class,
              "ExecutionPolicyInput",
              GridGrindContractText.executionPolicyInputSummary(),
              List.of("mode", "journal", "calculation")),
          plainTypeDescriptor(
              "calculationPolicyInputType",
              CalculationPolicyInput.class,
              "CalculationPolicyInput",
              GridGrindContractText.calculationPolicyInputSummary(),
              List.of("strategy")),
          plainTypeDescriptor(
              "executionModeInputType",
              ExecutionModeInput.class,
              "ExecutionModeInput",
              GridGrindContractText.executionModeInputSummary(),
              List.of("readMode", "writeMode")),
          plainTypeDescriptor(
              "executionJournalInputType",
              ExecutionJournalInput.class,
              "ExecutionJournalInput",
              GridGrindContractText.executionJournalInputSummary(),
              List.of("level")),
          plainTypeDescriptor(
              "calculationReportType",
              CalculationReport.class,
              "CalculationReport",
              "Structured calculation policy, preflight classification, and execution outcome"
                  + " returned on every success and failure response.",
              List.of("preflight")),
          plainTypeDescriptor(
              "calculationPreflightType",
              CalculationReport.Preflight.class,
              "CalculationPreflightReport",
              "Formula capability classification captured before server-side evaluation begins.",
              List.of()),
          plainTypeDescriptor(
              "calculationSummaryType",
              CalculationReport.Summary.class,
              "CalculationPreflightSummary",
              "Aggregate counts for evaluable, unevaluable, and unparseable formulas.",
              List.of()),
          plainTypeDescriptor(
              "formulaCapabilityType",
              CalculationReport.FormulaCapability.class,
              "FormulaCapabilityReport",
              "One classified formula-cell capability entry from calculation preflight.",
              List.of("problemCode", "message")),
          plainTypeDescriptor(
              "calculationExecutionType",
              CalculationReport.Execution.class,
              "CalculationExecutionReport",
              "Post-execution outcome for the authored calculation policy.",
              List.of("message")),
          plainTypeDescriptor(
              "formulaEnvironmentInputType",
              FormulaEnvironmentInput.class,
              "FormulaEnvironmentInput",
              "Request-scoped formula-evaluation environment covering external workbook bindings,"
                  + " missing-workbook policy, and template-backed UDF toolpacks.",
              List.of("externalWorkbooks", "missingWorkbookPolicy", "udfToolpacks")),
          plainTypeDescriptor(
              "ooxmlOpenSecurityInputType",
              OoxmlOpenSecurityInput.class,
              "OoxmlOpenSecurityInput",
              "Optional OOXML package-open settings for encrypted existing workbook sources."
                  + " password unlocks the encrypted OOXML package before GridGrind opens the"
                  + " inner .xlsx workbook.",
              List.of("password")),
          plainTypeDescriptor(
              "ooxmlPersistenceSecurityInputType",
              OoxmlPersistenceSecurityInput.class,
              "OoxmlPersistenceSecurityInput",
              "Optional OOXML package-security settings applied during persistence."
                  + " Supply encryption, signature, or both.",
              List.of("encryption", "signature")),
          plainTypeDescriptor(
              "ooxmlEncryptionInputType",
              OoxmlEncryptionInput.class,
              "OoxmlEncryptionInput",
              "OOXML package-encryption settings for workbook persistence."
                  + " mode defaults to AGILE when omitted.",
              List.of("mode")),
          plainTypeDescriptor(
              "ooxmlSignatureInputType",
              OoxmlSignatureInput.class,
              "OoxmlSignatureInput",
              "OOXML package-signing settings for workbook persistence."
                  + " pkcs12Path follows the request-owned path rule."
                  + " "
                  + GridGrindContractText.requestOwnedPathResolutionSummary()
                  + " keyPassword defaults to keystorePassword and digestAlgorithm defaults to"
                  + " SHA256 when omitted."
                  + " alias may be omitted only when the keystore contains exactly one"
                  + " signable private-key entry.",
              List.of("keyPassword", "alias", "digestAlgorithm", "description")),
          plainTypeDescriptor(
              "ooxmlPackageSecurityReportType",
              OoxmlPackageSecurityReport.class,
              "OoxmlPackageSecurityReport",
              "Factual OOXML package-security report covering encryption and package signatures.",
              List.of()),
          plainTypeDescriptor(
              "ooxmlEncryptionReportType",
              OoxmlEncryptionReport.class,
              "OoxmlEncryptionReport",
              "Factual OOXML package-encryption report for one workbook package."
                  + " Detail fields are present only when encrypted=true.",
              List.of()),
          plainTypeDescriptor(
              "ooxmlSignatureReportType",
              OoxmlSignatureReport.class,
              "OoxmlSignatureReport",
              "Factual OOXML package-signature report for one signature part."
                  + " state reflects the currently loaded workbook package, including"
                  + " INVALIDATED_BY_MUTATION for source signatures after in-memory edits.",
              List.of()),
          plainTypeDescriptor(
              "formulaExternalWorkbookInputType",
              FormulaExternalWorkbookInput.class,
              "FormulaExternalWorkbookInput",
              "One external workbook binding keyed by the workbook name used inside formulas."
                  + " path follows the request-owned path rule."
                  + " "
                  + GridGrindContractText.requestOwnedPathResolutionSummary(),
              List.of()),
          plainTypeDescriptor(
              "formulaUdfToolpackInputType",
              FormulaUdfToolpackInput.class,
              "FormulaUdfToolpackInput",
              "One named collection of template-backed user-defined functions.",
              List.of()),
          plainTypeDescriptor(
              "formulaUdfFunctionInputType",
              FormulaUdfFunctionInput.class,
              "FormulaUdfFunctionInput",
              "One template-backed user-defined function."
                  + " formulaTemplate may reference ARG1, ARG2, and higher placeholders."
                  + " maximumArgumentCount defaults to minimumArgumentCount when omitted.",
              List.of("maximumArgumentCount")),
          plainTypeDescriptor(
              "qualifiedCellAddressType",
              dev.erst.gridgrind.contract.selector.CellSelector.QualifiedAddress.class,
              "CellSelector.QualifiedAddress",
              "One workbook-qualified cell address used by selector-based targeted cell workflows.",
              List.of()),
          plainTypeDescriptor(
              "drawingMarkerInputType",
              DrawingMarkerInput.class,
              "DrawingMarkerInput",
              "One zero-based drawing marker with explicit column, row, and in-cell offsets.",
              List.of()),
          plainTypeDescriptor(
              "arrayFormulaInputType",
              ArrayFormulaInput.class,
              "ArrayFormulaInput",
              "One authored array formula bound to a contiguous single-cell or multi-cell range."
                  + " Leading = or {=...} wrappers normalize away for inline sources.",
              List.of()),
          plainTypeDescriptor(
              "customXmlMappingLocatorType",
              CustomXmlMappingLocator.class,
              "CustomXmlMappingLocator",
              "One locator for an existing workbook custom-XML mapping."
                  + " Supply mapId, name, or both; the locator must resolve to exactly one"
                  + " existing mapping.",
              List.of("mapId", "name")),
          plainTypeDescriptor(
              "customXmlImportInputType",
              CustomXmlImportInput.class,
              "CustomXmlImportInput",
              "One custom-XML import payload targeting an existing workbook mapping plus the XML"
                  + " content to import.",
              List.of()),
          plainTypeDescriptor(
              "chartInputType",
              ChartInput.class,
              "ChartInput",
              "One authored chart with a drawing anchor, chart-level presentation state,"
                  + " and one or more plots.",
              List.of("title", "legend", "displayBlanksAs", "plotOnlyVisibleCells")),
          plainTypeDescriptor(
              "chartAxisInputType",
              ChartInput.Axis.class,
              "ChartAxisInput",
              "One authored chart axis used by a chart plot."
                  + " visible defaults to true when omitted.",
              List.of("visible")),
          plainTypeDescriptor(
              "chartSeriesInputType",
              ChartInput.Series.class,
              "ChartSeriesInput",
              "One authored chart series with a title plus category and value data sources."
                  + " smooth, marker, and explosion fields are optional by chart family.",
              List.of("title", "smooth", "markerStyle", "markerSize", "explosion")),
          plainTypeDescriptor(
              "pictureDataInputType",
              PictureDataInput.class,
              "PictureDataInput",
              "One picture payload with explicit format and base64-encoded binary data.",
              List.of()),
          plainTypeDescriptor(
              "pictureInputType",
              PictureInput.class,
              "PictureInput",
              "Named picture-authoring payload for SET_PICTURE.",
              List.of("description")),
          plainTypeDescriptor(
              "signatureLineInputType",
              SignatureLineInput.class,
              "SignatureLineInput",
              "Named signature-line authoring payload for SET_SIGNATURE_LINE."
                  + " allowComments defaults to true when omitted and plainSignature is optional,"
                  + " but caption or suggested signer metadata must still be present.",
              List.of(
                  "allowComments",
                  "signingInstructions",
                  "suggestedSigner",
                  "suggestedSigner2",
                  "suggestedSignerEmail",
                  "caption",
                  "invalidStamp",
                  "plainSignature")),
          plainTypeDescriptor(
              "shapeInputType",
              ShapeInput.class,
              "ShapeInput",
              "Named simple-shape or connector authoring payload for SET_SHAPE."
                  + " kind is limited to the authored drawing shape family."
                  + " presetGeometryToken defaults to rect for SIMPLE_SHAPE when omitted."
                  + " Invalid presetGeometryToken values are rejected non-mutatingly.",
              List.of("presetGeometryToken", "text")),
          plainTypeDescriptor(
              "embeddedObjectInputType",
              EmbeddedObjectInput.class,
              "EmbeddedObjectInput",
              "Named embedded-object authoring payload for SET_EMBEDDED_OBJECT."
                  + " base64Data holds the embedded package bytes and previewImage holds the"
                  + " visible preview raster.",
              List.of()),
          plainTypeDescriptor(
              "commentInputType",
              CommentInput.class,
              "CommentInput",
              "Comment payload attached to one cell."
                  + " Comments can carry ordered rich-text runs and an explicit anchor box.",
              List.of("visible", "runs", "anchor")),
          plainTypeDescriptor(
              "commentAnchorInputType",
              CommentAnchorInput.class,
              "CommentAnchorInput",
              "Explicit comment-anchor bounds measured in zero-based column and row indexes.",
              List.of()),
          plainTypeDescriptor(
              "namedRangeTargetType",
              NamedRangeTarget.class,
              "NamedRangeTarget",
              "Named-range target payload."
                  + " Supply either sheetName plus range, or formula by itself.",
              List.of("sheetName", "range", "formula")),
          plainTypeDescriptor(
              "sheetProtectionSettingsType",
              SheetProtectionSettings.class,
              "SheetProtectionSettings",
              "Supported sheet-protection lock flags authored and reported by GridGrind.",
              List.of()),
          plainTypeDescriptor(
              "cellStyleInputType",
              CellStyleInput.class,
              "CellStyleInput",
              "Style patch applied to a cell or range; at least one field must be set."
                  + " Colors use #RRGGBB hex and style subgroups are nested explicitly.",
              List.of("numberFormat", "alignment", "font", "fill", "border", "protection")),
          plainTypeDescriptor(
              "cellAlignmentInputType",
              CellAlignmentInput.class,
              "CellAlignmentInput",
              "Alignment patch for cell styling; at least one field must be set."
                  + " textRotation uses XSSF's explicit 0-180 degree scale and indentation uses"
                  + " Excel's 0-250 cell-indent range.",
              List.of(
                  "wrapText",
                  "horizontalAlignment",
                  "verticalAlignment",
                  "textRotation",
                  "indentation")),
          plainTypeDescriptor(
              "cellFontInputType",
              CellFontInput.class,
              "CellFontInput",
              "Font patch for cell styling; at least one field must be set."
                  + " Colors can use RGB, theme, indexed, and tint semantics.",
              List.of(
                  "bold",
                  "italic",
                  "fontName",
                  "fontHeight",
                  "fontColor",
                  "fontColorTheme",
                  "fontColorIndexed",
                  "fontColorTint",
                  "underline",
                  "strikeout")),
          plainTypeDescriptor(
              "colorInputType",
              ColorInput.class,
              "ColorInput",
              "Color payload preserving RGB, theme, indexed, and tint semantics."
                  + " At least one of rgb, theme, or indexed must be supplied.",
              List.of("rgb", "theme", "indexed", "tint")),
          plainTypeDescriptor(
              "richTextRunInputType",
              RichTextRunInput.class,
              "RichTextRunInput",
              "One ordered rich-text run for a string cell."
                  + " text must be non-empty; font is an optional override patch."
                  + " The ordered run texts concatenate to the stored plain string value.",
              List.of("font")),
          plainTypeDescriptor(
              "cellFillInputType",
              CellFillInput.class,
              "CellFillInput",
              "Fill patch for cell styling. pattern controls solid and patterned fills;"
                  + " colors can use RGB, theme, indexed, and tint semantics."
                  + " gradient is mutually exclusive with patterned fill fields.",
              List.of(
                  "pattern",
                  "foregroundColor",
                  "foregroundColorTheme",
                  "foregroundColorIndexed",
                  "foregroundColorTint",
                  "backgroundColor",
                  "backgroundColorTheme",
                  "backgroundColorIndexed",
                  "backgroundColorTint",
                  "gradient")),
          plainTypeDescriptor(
              "cellGradientFillInputType",
              CellGradientFillInput.class,
              "CellGradientFillInput",
              "Gradient fill payload for cell-style authoring."
                  + " LINEAR gradients use degree, PATH gradients use left/right/top/bottom,"
                  + " and the two geometry modes must not be mixed."
                  + " stops must contain at least two entries.",
              List.of("type", "degree", "left", "right", "top", "bottom")),
          plainTypeDescriptor(
              "cellGradientStopInputType",
              CellGradientStopInput.class,
              "CellGradientStopInput",
              "One gradient stop with a normalized position between 0.0 and 1.0.",
              List.of()),
          plainTypeDescriptor(
              "cellBorderInputType",
              CellBorderInput.class,
              "CellBorderInput",
              "Border patch for cell styling; at least one side must be set."
                  + " Use 'all' as shorthand for all four sides.",
              List.of("all", "top", "right", "bottom", "left")),
          plainTypeDescriptor(
              "cellBorderSideInputType",
              CellBorderSideInput.class,
              "CellBorderSideInput",
              "One border side defined by its border style and optional color semantics.",
              List.of("style", "color", "colorTheme", "colorIndexed", "colorTint")),
          plainTypeDescriptor(
              "cellProtectionInputType",
              CellProtectionInput.class,
              "CellProtectionInput",
              "Cell protection patch; at least one field must be set."
                  + " These flags matter when sheet protection is enabled.",
              List.of("locked", "hiddenFormula")),
          plainTypeDescriptor(
              "dataValidationInputType",
              DataValidationInput.class,
              "DataValidationInput",
              "Supported data-validation definition attached to one sheet range.",
              List.of("allowBlank", "suppressDropDownArrow", "prompt", "errorAlert")),
          plainTypeDescriptor(
              "dataValidationPromptInputType",
              DataValidationPromptInput.class,
              "DataValidationPromptInput",
              "Optional prompt-box configuration shown when a validated cell is selected.",
              List.of("showPromptBox")),
          plainTypeDescriptor(
              "dataValidationErrorAlertInputType",
              DataValidationErrorAlertInput.class,
              "DataValidationErrorAlertInput",
              "Optional error-box configuration shown when invalid data is entered.",
              List.of("showErrorBox")),
          plainTypeDescriptor(
              "autofilterCustomConditionInputType",
              AutofilterFilterCriterionInput.CustomConditionInput.class,
              "AutofilterCustomConditionInput",
              "One comparator-value pair nested inside a custom autofilter criterion.",
              List.of()),
          plainTypeDescriptor(
              "autofilterFilterColumnInputType",
              AutofilterFilterColumnInput.class,
              "AutofilterFilterColumnInput",
              "One authored autofilter filter-column payload with an explicit column criterion.",
              List.of("showButton")),
          plainTypeDescriptor(
              "autofilterSortConditionInputType",
              AutofilterSortConditionInput.class,
              "AutofilterSortConditionInput",
              "One authored sort condition nested inside an autofilter sort state.",
              List.of("sortBy", "color", "iconId")),
          plainTypeDescriptor(
              "autofilterSortStateInputType",
              AutofilterSortStateInput.class,
              "AutofilterSortStateInput",
              "Authored autofilter sort-state payload with one or more ordered sort conditions.",
              List.of("caseSensitive", "columnSort", "sortMethod")),
          plainTypeDescriptor(
              "conditionalFormattingBlockInputType",
              ConditionalFormattingBlockInput.class,
              "ConditionalFormattingBlockInput",
              "One authored conditional-formatting block with ordered target ranges and rules."
                  + " rules must not be empty; ranges must be unique.",
              List.of()),
          plainTypeDescriptor(
              "conditionalFormattingThresholdInputType",
              ConditionalFormattingThresholdInput.class,
              "ConditionalFormattingThresholdInput",
              "Threshold payload shared by authored advanced conditional-formatting rules.",
              List.of("formula", "value")),
          plainTypeDescriptor(
              "headerFooterTextInputType",
              HeaderFooterTextInput.class,
              "HeaderFooterTextInput",
              "Plain left, center, and right header or footer text segments."
                  + " Null fields default to empty string.",
              List.of("left", "center", "right")),
          plainTypeDescriptor(
              "differentialStyleInputType",
              DifferentialStyleInput.class,
              "DifferentialStyleInput",
              "Differential style payload used by authored conditional-formatting rules."
                  + " At least one field must be set. Colors use #RRGGBB hex.",
              List.of(
                  "numberFormat",
                  "bold",
                  "italic",
                  "fontHeight",
                  "fontColor",
                  "underline",
                  "strikeout",
                  "fillColor",
                  "border")),
          plainTypeDescriptor(
              "differentialBorderInputType",
              DifferentialBorderInput.class,
              "DifferentialBorderInput",
              "Conditional-formatting differential border patch; at least one side must be set."
                  + " Use 'all' as shorthand for all four sides.",
              List.of("all", "top", "right", "bottom", "left")),
          plainTypeDescriptor(
              "differentialBorderSideInputType",
              DifferentialBorderSideInput.class,
              "DifferentialBorderSideInput",
              "One conditional-formatting differential border side defined by style and optional"
                  + " color.",
              List.of()),
          plainTypeDescriptor(
              "ignoredErrorInputType",
              IgnoredErrorInput.class,
              "IgnoredErrorInput",
              "One ignored-error block anchored to one A1-style range plus one or more"
                  + " ignored-error families.",
              List.of()),
          plainTypeDescriptor(
              "printLayoutInputType",
              PrintLayoutInput.class,
              "PrintLayoutInput",
              "Authoritative supported print-layout payload for one SET_PRINT_LAYOUT request."
                  + " All fields are optional and normalize to defaults when omitted.",
              List.of(
                  "printArea",
                  "orientation",
                  "scaling",
                  "repeatingRows",
                  "repeatingColumns",
                  "header",
                  "footer",
                  "setup")),
          plainTypeDescriptor(
              "printMarginsInputType",
              PrintMarginsInput.class,
              "PrintMarginsInput",
              "Explicit print margins measured in the workbook's stored inch-based values.",
              List.of()),
          plainTypeDescriptor(
              "printSetupInputType",
              PrintSetupInput.class,
              "PrintSetupInput",
              "Advanced page-setup payload nested under print-layout authoring."
                  + " All fields are optional and normalize to defaults when omitted.",
              List.of(
                  "margins",
                  "printGridlines",
                  "horizontallyCentered",
                  "verticallyCentered",
                  "paperSize",
                  "draft",
                  "blackAndWhite",
                  "copies",
                  "useFirstPageNumber",
                  "firstPageNumber",
                  "rowBreaks",
                  "columnBreaks")),
          plainTypeDescriptor(
              "sheetDefaultsInputType",
              SheetDefaultsInput.class,
              "SheetDefaultsInput",
              "Default row and column sizing authored as part of sheet-presentation state."
                  + " All fields are optional and normalize to defaults when omitted."
                  + " defaultColumnWidth must be > 0 and <= 255;"
                  + " defaultRowHeightPoints must be > 0 and <= "
                  + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS
                  + ".",
              List.of("defaultColumnWidth", "defaultRowHeightPoints")),
          plainTypeDescriptor(
              "sheetDisplayInputType",
              SheetDisplayInput.class,
              "SheetDisplayInput",
              "Screen-facing sheet display flags authored as part of sheet-presentation state."
                  + " All fields are optional and normalize to defaults when omitted.",
              List.of(
                  "displayGridlines",
                  "displayZeros",
                  "displayRowColHeadings",
                  "displayFormulas",
                  "rightToLeft")),
          plainTypeDescriptor(
              "sheetOutlineSummaryInputType",
              SheetOutlineSummaryInput.class,
              "SheetOutlineSummaryInput",
              "Outline-summary placement authored as part of sheet-presentation state."
                  + " All fields are optional and normalize to defaults when omitted.",
              List.of("rowSumsBelow", "rowSumsRight")),
          plainTypeDescriptor(
              "sheetPresentationInputType",
              SheetPresentationInput.class,
              "SheetPresentationInput",
              "Authoritative sheet-presentation payload for one SET_SHEET_PRESENTATION request."
                  + " All fields are optional and normalize to defaults or clear state when"
                  + " omitted.",
              List.of("display", "tabColor", "outlineSummary", "sheetDefaults", "ignoredErrors")),
          plainTypeDescriptor(
              "pivotTableInputType",
              PivotTableInput.class,
              "PivotTableInput",
              "Workbook-global pivot-table definition for one SET_PIVOT_TABLE request."
                  + " Source-column assignments across rowLabels, columnLabels, reportFilters,"
                  + " and dataFields must be disjoint."
                  + " reportFilters require anchor.topLeftAddress on row 3 or lower.",
              List.of("rowLabels", "columnLabels", "reportFilters")),
          plainTypeDescriptor(
              "pivotTableAnchorInputType",
              PivotTableInput.Anchor.class,
              "PivotTableAnchorInput",
              "Top-left anchor for a pivot table rendered on its destination sheet."
                  + " The address must be a single-cell A1 reference.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableDataFieldInputType",
              PivotTableInput.DataField.class,
              "PivotTableDataFieldInput",
              "One authored pivot data field bound to a source column and aggregation function."
                  + " displayName defaults to sourceColumnName when omitted.",
              List.of("displayName", "valueFormat")),
          plainTypeDescriptor(
              "tableColumnInputType",
              TableColumnInput.class,
              "TableColumnInput",
              "Advanced table-column metadata applied by zero-based ordinal column index.",
              List.of(
                  "uniqueName", "totalsRowLabel", "totalsRowFunction", "calculatedColumnFormula")),
          plainTypeDescriptor(
              "tableInputType",
              TableInput.class,
              "TableInput",
              "Workbook-global table definition for one SET_TABLE request.",
              List.of(
                  "showTotalsRow",
                  "hasAutofilter",
                  "comment",
                  "published",
                  "insertRow",
                  "insertRowShift",
                  "headerRowCellStyle",
                  "dataCellStyle",
                  "totalsRowCellStyle",
                  "columns")),
          plainTypeDescriptor(
              "workbookProtectionInputType",
              WorkbookProtectionInput.class,
              "WorkbookProtectionInput",
              "Workbook-protection payload covering workbook and revisions lock state plus"
                  + " optional passwords.",
              List.of(
                  "structureLocked",
                  "windowsLocked",
                  "revisionsLocked",
                  "workbookPassword",
                  "revisionsPassword")),
          plainTypeDescriptor(
              "workbookProtectionReportType",
              WorkbookProtectionReport.class,
              "WorkbookProtectionReport",
              "Exact workbook-protection report covering structure, windows, revisions,"
                  + " and password-hash presence flags.",
              List.of()),
          plainTypeDescriptor(
              "sheetSummaryReportType",
              GridGrindResponse.SheetSummaryReport.class,
              "SheetSummaryReport",
              "Exact sheet summary report including visibility, protection, and structural counts.",
              List.of()),
          plainTypeDescriptor(
              "cellStyleReportType",
              GridGrindResponse.CellStyleReport.class,
              "CellStyleReport",
              "Exact effective cell-style report used by style assertions.",
              List.of()),
          plainTypeDescriptor(
              "cellAlignmentReportType",
              CellAlignmentReport.class,
              "CellAlignmentReport",
              "Exact cell-alignment report.",
              List.of()),
          plainTypeDescriptor(
              "cellFontReportType",
              CellFontReport.class,
              "CellFontReport",
              "Exact cell-font report.",
              List.of("fontColor")),
          plainTypeDescriptor(
              "cellFillReportType",
              CellFillReport.class,
              "CellFillReport",
              "Exact cell-fill report including pattern, colors, or gradient payload.",
              List.of("foregroundColor", "backgroundColor", "gradient")),
          plainTypeDescriptor(
              "cellBorderReportType",
              CellBorderReport.class,
              "CellBorderReport",
              "Exact four-sided cell-border report.",
              List.of()),
          plainTypeDescriptor(
              "cellBorderSideReportType",
              CellBorderSideReport.class,
              "CellBorderSideReport",
              "Exact one-sided cell-border report.",
              List.of("color")),
          plainTypeDescriptor(
              "cellProtectionReportType",
              CellProtectionReport.class,
              "CellProtectionReport",
              "Exact cell-protection report.",
              List.of()),
          plainTypeDescriptor(
              "cellColorReportType",
              CellColorReport.class,
              "CellColorReport",
              "Exact workbook color report preserving RGB, theme, indexed, and tint semantics.",
              List.of("rgb", "theme", "indexed", "tint")),
          plainTypeDescriptor(
              "fontHeightReportType",
              FontHeightReport.class,
              "FontHeightReport",
              "Exact font-height report expressed in twips and points.",
              List.of()),
          plainTypeDescriptor(
              "cellGradientFillReportType",
              CellGradientFillReport.class,
              "CellGradientFillReport",
              "Exact gradient-fill report with geometry and stops.",
              List.of("degree", "left", "right", "top", "bottom")),
          plainTypeDescriptor(
              "cellGradientStopReportType",
              CellGradientStopReport.class,
              "CellGradientStopReport",
              "Exact gradient stop report.",
              List.of()),
          plainTypeDescriptor(
              "tableEntryReportType",
              TableEntryReport.class,
              "TableEntryReport",
              "Exact workbook table report used by table-facts assertions.",
              List.of("comment", "headerRowCellStyle", "dataCellStyle", "totalsRowCellStyle")),
          plainTypeDescriptor(
              "tableColumnReportType",
              TableColumnReport.class,
              "TableColumnReport",
              "Exact table-column report.",
              List.of(
                  "uniqueName", "totalsRowLabel", "totalsRowFunction", "calculatedColumnFormula")),
          plainTypeDescriptor(
              "drawingMarkerReportType",
              DrawingMarkerReport.class,
              "DrawingMarkerReport",
              "Exact cell-relative drawing marker report.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableAnchorReportType",
              PivotTableReport.Anchor.class,
              "PivotTableAnchorReport",
              "Exact pivot-table anchor report.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableFieldReportType",
              PivotTableReport.Field.class,
              "PivotTableFieldReport",
              "Exact pivot field report bound to one source column.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableDataFieldReportType",
              PivotTableReport.DataField.class,
              "PivotTableDataFieldReport",
              "Exact pivot data-field report.",
              List.of("valueFormat")),
          plainTypeDescriptor(
              "arrayFormulaReportType",
              ArrayFormulaReport.class,
              "ArrayFormulaReport",
              "One factual array-formula group report returned by GET_ARRAY_FORMULAS.",
              List.of()),
          plainTypeDescriptor(
              "customXmlMappingReportType",
              CustomXmlMappingReport.class,
              "CustomXmlMappingReport",
              "One factual workbook custom-XML mapping report.",
              List.of(
                  "schemaNamespace",
                  "schemaLanguage",
                  "schemaReference",
                  "schemaXml",
                  "dataBinding")),
          plainTypeDescriptor(
              "customXmlDataBindingReportType",
              CustomXmlDataBindingReport.class,
              "CustomXmlDataBindingReport",
              "Optional custom-XML data-binding metadata attached to one workbook mapping.",
              List.of("dataBindingName", "fileBinding", "connectionId", "fileBindingName")),
          plainTypeDescriptor(
              "customXmlLinkedCellReportType",
              CustomXmlLinkedCellReport.class,
              "CustomXmlLinkedCellReport",
              "One single-cell binding linked to a custom-XML mapping.",
              List.of()),
          plainTypeDescriptor(
              "customXmlLinkedTableReportType",
              CustomXmlLinkedTableReport.class,
              "CustomXmlLinkedTableReport",
              "One XML-mapped table linked to a custom-XML mapping.",
              List.of()),
          plainTypeDescriptor(
              "customXmlExportReportType",
              CustomXmlExportReport.class,
              "CustomXmlExportReport",
              "One exported custom-XML mapping payload plus the factual mapping metadata used to"
                  + " produce it.",
              List.of()),
          plainTypeDescriptor(
              "chartReportType",
              ChartReport.class,
              "ChartReport",
              "One factual chart report with chart-level presentation state and one or more"
                  + " plots.",
              List.of()),
          plainTypeDescriptor(
              "chartAxisReportType",
              ChartReport.Axis.class,
              "ChartAxisReport",
              "Exact chart-axis report.",
              List.of()),
          plainTypeDescriptor(
              "chartSeriesReportType",
              ChartReport.Series.class,
              "ChartSeriesReport",
              "Exact chart-series report."
                  + " smooth, marker, and explosion fields are populated only when the"
                  + " stored plot family supports them.",
              List.of("smooth", "markerStyle", "markerSize", "explosion")));

  private static CatalogNestedTypeDescriptor nestedTypeGroup(
      String group, Class<?> sealedType, List<CatalogTypeDescriptor> typeDescriptors) {
    return GridGrindProtocolCatalog.nestedTypeGroup(group, sealedType, typeDescriptors);
  }

  private static CatalogPlainTypeDescriptor plainTypeDescriptor(
      String group,
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> optionalFields) {
    return GridGrindProtocolCatalog.plainTypeDescriptor(
        group, recordType, id, summary, optionalFields);
  }

  private static CatalogTypeDescriptor descriptor(
      Class<? extends Record> recordType, String id, String summary, String... optionalFields) {
    return GridGrindProtocolCatalog.descriptor(recordType, id, summary, optionalFields);
  }
}
