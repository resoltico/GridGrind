package dev.erst.gridgrind.jazzer.support;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelDrawingShapeKind;
import dev.erst.gridgrind.excel.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.excel.ExcelFillPattern;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelPictureFormat;
import dev.erst.gridgrind.excel.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.ExcelPivotTableDefinition;
import dev.erst.gridgrind.excel.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.protocol.dto.AnalysisFindingCode;
import dev.erst.gridgrind.protocol.dto.AnalysisSeverity;
import dev.erst.gridgrind.protocol.dto.AutofilterEntryReport;
import dev.erst.gridgrind.protocol.dto.AutofilterFilterColumnReport;
import dev.erst.gridgrind.protocol.dto.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.protocol.dto.AutofilterHealthReport;
import dev.erst.gridgrind.protocol.dto.AutofilterSortConditionReport;
import dev.erst.gridgrind.protocol.dto.AutofilterSortStateReport;
import dev.erst.gridgrind.protocol.dto.CellAlignmentReport;
import dev.erst.gridgrind.protocol.dto.CellBorderReport;
import dev.erst.gridgrind.protocol.dto.CellBorderSideReport;
import dev.erst.gridgrind.protocol.dto.CellColorReport;
import dev.erst.gridgrind.protocol.dto.CellFillReport;
import dev.erst.gridgrind.protocol.dto.CellFontReport;
import dev.erst.gridgrind.protocol.dto.CellGradientFillReport;
import dev.erst.gridgrind.protocol.dto.CellGradientStopReport;
import dev.erst.gridgrind.protocol.dto.CellProtectionReport;
import dev.erst.gridgrind.protocol.dto.CellSelection;
import dev.erst.gridgrind.protocol.dto.ChartReport;
import dev.erst.gridgrind.protocol.dto.CommentAnchorReport;
import dev.erst.gridgrind.protocol.dto.DataValidationEntryReport;
import dev.erst.gridgrind.protocol.dto.DataValidationHealthReport;
import dev.erst.gridgrind.protocol.dto.DataValidationRuleInput;
import dev.erst.gridgrind.protocol.dto.DrawingAnchorReport;
import dev.erst.gridgrind.protocol.dto.DrawingMarkerReport;
import dev.erst.gridgrind.protocol.dto.DrawingObjectPayloadReport;
import dev.erst.gridgrind.protocol.dto.DrawingObjectReport;
import dev.erst.gridgrind.protocol.dto.FontHeightReport;
import dev.erst.gridgrind.protocol.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.protocol.dto.GridGrindRequest;
import dev.erst.gridgrind.protocol.dto.GridGrindResponse;
import dev.erst.gridgrind.protocol.dto.HeaderFooterTextReport;
import dev.erst.gridgrind.protocol.dto.HyperlinkTarget;
import dev.erst.gridgrind.protocol.dto.NamedRangeScope;
import dev.erst.gridgrind.protocol.dto.NamedRangeSelection;
import dev.erst.gridgrind.protocol.dto.NamedRangeTarget;
import dev.erst.gridgrind.protocol.dto.PaneReport;
import dev.erst.gridgrind.protocol.dto.PivotTableHealthReport;
import dev.erst.gridgrind.protocol.dto.PivotTableReport;
import dev.erst.gridgrind.protocol.dto.PivotTableSelection;
import dev.erst.gridgrind.protocol.dto.PrintAreaReport;
import dev.erst.gridgrind.protocol.dto.PrintLayoutReport;
import dev.erst.gridgrind.protocol.dto.PrintMarginsReport;
import dev.erst.gridgrind.protocol.dto.PrintScalingReport;
import dev.erst.gridgrind.protocol.dto.PrintSetupReport;
import dev.erst.gridgrind.protocol.dto.PrintTitleColumnsReport;
import dev.erst.gridgrind.protocol.dto.PrintTitleRowsReport;
import dev.erst.gridgrind.protocol.dto.RangeSelection;
import dev.erst.gridgrind.protocol.dto.RequestWarning;
import dev.erst.gridgrind.protocol.dto.RichTextRunReport;
import dev.erst.gridgrind.protocol.dto.SheetProtectionSettings;
import dev.erst.gridgrind.protocol.dto.SheetSelection;
import dev.erst.gridgrind.protocol.dto.TableColumnReport;
import dev.erst.gridgrind.protocol.dto.TableEntryReport;
import dev.erst.gridgrind.protocol.dto.TableHealthReport;
import dev.erst.gridgrind.protocol.dto.TableSelection;
import dev.erst.gridgrind.protocol.dto.TableStyleReport;
import dev.erst.gridgrind.protocol.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Tests for WorkbookInvariantChecks success-shape validation. */
class WorkbookInvariantChecksTest {
  @Test
  void acceptsSuccessResponsesWithOrderedReads(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");
    GridGrindResponse.CellStyleReport style = defaultStyle();

    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                "result.xlsx", workbookPath.toString()),
            List.of(
                new RequestWarning(
                    1, "SET_CELL", "Formula references same-request sheet names with spaces.")),
            List.of(
                new WorkbookReadResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of("Budget"), 1, false)),
                new WorkbookReadResult.NamedRangesResult(
                    "ranges",
                    List.of(
                        new GridGrindResponse.NamedRangeReport.RangeReport(
                            "BudgetTotal",
                            new NamedRangeScope.Workbook(),
                            "Budget!$B$4",
                            new NamedRangeTarget("Budget", "B4")))),
                new WorkbookReadResult.CellsResult(
                    "cells",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellReport.TextReport(
                            "A1",
                            "STRING",
                            "Report",
                            style,
                            new HyperlinkTarget.Url("https://example.com/report"),
                            new GridGrindResponse.CommentReport("Review", "GridGrind", true),
                            "Report",
                            List.of(
                                new RichTextRunReport(
                                    "Report",
                                    new CellFontReport(
                                        false,
                                        false,
                                        "Calibri",
                                        new FontHeightReport(220, new BigDecimal("11")),
                                        null,
                                        false,
                                        false)))))),
                new WorkbookReadResult.WindowResult(
                    "window",
                    new GridGrindResponse.WindowReport(
                        "Budget",
                        "A1",
                        1,
                        1,
                        List.of(
                            new GridGrindResponse.WindowRowReport(
                                0,
                                List.of(
                                    new GridGrindResponse.CellReport.TextReport(
                                        "A1", "STRING", "Report", style, null, null, "Report",
                                        null)))))),
                new WorkbookReadResult.MergedRegionsResult(
                    "merged", "Budget", List.of(new GridGrindResponse.MergedRegionReport("A1:B1"))),
                new WorkbookReadResult.HyperlinksResult(
                    "hyperlinks",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellHyperlinkReport(
                            "A1", new HyperlinkTarget.Url("https://example.com/report")))),
                new WorkbookReadResult.CommentsResult(
                    "comments",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellCommentReport(
                            "A1",
                            new GridGrindResponse.CommentReport("Review", "GridGrind", true)))),
                new WorkbookReadResult.SheetLayoutResult(
                    "layout",
                    new GridGrindResponse.SheetLayoutReport(
                        "Budget",
                        new PaneReport.Frozen(1, 1, 1, 1),
                        125,
                        dev.erst.gridgrind.protocol.dto.SheetPresentationReport.defaults(),
                        List.of(new GridGrindResponse.ColumnLayoutReport(0, 12.5, false, 0, false)),
                        List.of(new GridGrindResponse.RowLayoutReport(0, 18.0, false, 0, false)))),
                new WorkbookReadResult.PrintLayoutResult(
                    "print-layout",
                    new PrintLayoutReport(
                        "Budget",
                        new PrintAreaReport.Range("A1:B20"),
                        ExcelPrintOrientation.LANDSCAPE,
                        new PrintScalingReport.Fit(1, 0),
                        new PrintTitleRowsReport.Band(0, 0),
                        new PrintTitleColumnsReport.Band(0, 0),
                        new HeaderFooterTextReport("Budget", "", ""),
                        new HeaderFooterTextReport("", "Page &P", ""))),
                new WorkbookReadResult.DataValidationsResult(
                    "data-validations",
                    "Budget",
                    List.of(
                        new DataValidationEntryReport.Supported(
                            List.of("A2:A5"),
                            new DataValidationEntryReport.DataValidationDefinitionReport(
                                new DataValidationRuleInput.WholeNumber(
                                    ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null),
                                true,
                                false,
                                null,
                                null)))),
                new WorkbookReadResult.FormulaSurfaceResult(
                    "formula-surface",
                    new GridGrindResponse.FormulaSurfaceReport(
                        1,
                        List.of(
                            new GridGrindResponse.SheetFormulaSurfaceReport(
                                "Budget",
                                1,
                                1,
                                List.of(
                                    new GridGrindResponse.FormulaPatternReport(
                                        "SUM(B2:B3)", 1, List.of("B4"))))))),
                new WorkbookReadResult.SheetSchemaResult(
                    "schema",
                    new GridGrindResponse.SheetSchemaReport(
                        "Budget",
                        "A1",
                        2,
                        1,
                        1,
                        List.of(
                            new GridGrindResponse.SchemaColumnReport(
                                0,
                                "A1",
                                "Item",
                                1,
                                0,
                                List.of(new GridGrindResponse.TypeCountReport("STRING", 1)),
                                "STRING")))),
                new WorkbookReadResult.NamedRangeSurfaceResult(
                    "named-range-surface",
                    new GridGrindResponse.NamedRangeSurfaceReport(
                        1,
                        0,
                        1,
                        0,
                        List.of(
                            new GridGrindResponse.NamedRangeSurfaceEntryReport(
                                "BudgetTotal",
                                new NamedRangeScope.Workbook(),
                                "Budget!$B$4",
                                GridGrindResponse.NamedRangeBackingKind.RANGE)))),
                new WorkbookReadResult.FormulaHealthResult(
                    "formula-health",
                    new GridGrindResponse.FormulaHealthReport(
                        1,
                        new GridGrindResponse.AnalysisSummaryReport(1, 0, 0, 1),
                        List.of(
                            new GridGrindResponse.AnalysisFindingReport(
                                AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                                AnalysisSeverity.INFO,
                                "Volatile formula",
                                "Formula uses NOW().",
                                new GridGrindResponse.AnalysisLocationReport.Cell("Budget", "B4"),
                                List.of("NOW()"))))),
                new WorkbookReadResult.DataValidationHealthResult(
                    "data-validation-health",
                    new DataValidationHealthReport(
                        1, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of()))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsWorkflowOutcomeShapeForOrderedReads(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadOperation.GetWorkbookSummary("summary"),
                new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1")),
                new WorkbookReadOperation.GetDataValidations(
                    "data-validations", "Budget", new RangeSelection.All()),
                new WorkbookReadOperation.GetHyperlinks(
                    "hyperlinks", "Budget", new CellSelection.AllUsedCells()),
                new WorkbookReadOperation.AnalyzeDataValidationHealth(
                    "data-validation-health", new SheetSelection.All()),
                new WorkbookReadOperation.AnalyzeNamedRangeHealth(
                    "named-range-health", new NamedRangeSelection.All())));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                workbookPath.toString(), workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of("Budget"), 1, false)),
                new WorkbookReadResult.CellsResult(
                    "cells", "Budget", List.of(textCell("A1", "Report"))),
                new WorkbookReadResult.DataValidationsResult(
                    "data-validations",
                    "Budget",
                    List.of(
                        new DataValidationEntryReport.Unsupported(
                            List.of("A2:A5"), "MISSING_FORMULA", "Formula is missing"))),
                new WorkbookReadResult.HyperlinksResult(
                    "hyperlinks",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellHyperlinkReport(
                            "A1", new HyperlinkTarget.Url("https://example.com/report")))),
                new WorkbookReadResult.DataValidationHealthResult(
                    "data-validation-health",
                    new DataValidationHealthReport(
                        1,
                        new GridGrindResponse.AnalysisSummaryReport(1, 1, 0, 0),
                        List.of(
                            new GridGrindResponse.AnalysisFindingReport(
                                AnalysisFindingCode.DATA_VALIDATION_UNSUPPORTED_RULE,
                                AnalysisSeverity.WARNING,
                                "Unsupported data-validation rule",
                                "Formula is missing",
                                new GridGrindResponse.AnalysisLocationReport.Range(
                                    "Budget", "A2:A5"),
                                List.of("A2:A5"))))),
                new WorkbookReadResult.NamedRangeHealthResult(
                    "named-range-health",
                    new GridGrindResponse.NamedRangeHealthReport(
                        1, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of()))));

    assertDoesNotThrow(
        () -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }

  @Test
  void acceptsNormalizedFileHyperlinkTargets(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                workbookPath.toString(), workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadResult.HyperlinksResult(
                    "hyperlinks",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellHyperlinkReport(
                            "A1", new HyperlinkTarget.File("/tmp/report.xlsx"))))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsWorkbookShapeWithNamedRanges() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("B4", ExcelCellValue.number(61.0));
      workbook.setNamedRange(
          new dev.erst.gridgrind.excel.ExcelNamedRangeDefinition(
              "BudgetTotal",
              new dev.erst.gridgrind.excel.ExcelNamedRangeScope.WorkbookScope(),
              new dev.erst.gridgrind.excel.ExcelNamedRangeTarget("Budget", "B4")));

      assertDoesNotThrow(() -> WorkbookInvariantChecks.requireWorkbookShape(workbook));
    }
  }

  @Test
  void acceptsWorkbookShapeWithB1SheetState() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      workbook.setActiveSheet("Beta");
      workbook.setSelectedSheets(List.of("Alpha", "Beta"));
      workbook.setSheetVisibility("Alpha", ExcelSheetVisibility.HIDDEN);
      workbook.setSheetProtection("Beta", protectionSettings());

      assertDoesNotThrow(() -> WorkbookInvariantChecks.requireWorkbookShape(workbook));
    }
  }

  @Test
  void acceptsResponseShapeWithProtectedSheetSummary(@TempDir Path tempDirectory)
      throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                "result.xlsx", workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadResult.SheetSummaryResult(
                    "sheet",
                    new GridGrindResponse.SheetSummaryReport(
                        "Budget",
                        ExcelSheetVisibility.VERY_HIDDEN,
                        new GridGrindResponse.SheetProtectionReport.Protected(
                            protocolProtectionSettings()),
                        4,
                        7,
                        3))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsSuccessResponsesWithAutofilterAndTableReads(@TempDir Path tempDirectory)
      throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                "result.xlsx", workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadResult.AutofiltersResult(
                    "autofilters",
                    "Budget",
                    List.of(
                        new AutofilterEntryReport.SheetOwned("E1:F3"),
                        new AutofilterEntryReport.TableOwned("A1:C4", "BudgetTable"))),
                new WorkbookReadResult.TablesResult(
                    "tables",
                    List.of(
                        new TableEntryReport(
                            "BudgetTable",
                            "Budget",
                            "A1:C4",
                            1,
                            0,
                            List.of("Item", "Amount", "Billable"),
                            new TableStyleReport.Named(
                                "TableStyleMedium2", false, false, true, false),
                            true))),
                new WorkbookReadResult.AutofilterHealthResult(
                    "autofilter-health",
                    new AutofilterHealthReport(
                        2, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of())),
                new WorkbookReadResult.TableHealthResult(
                    "table-health",
                    new TableHealthReport(
                        1, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of()))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsWorkflowOutcomeShapeForAutofilterAndTableReads(@TempDir Path tempDirectory)
      throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadOperation.GetAutofilters("autofilters", "Budget"),
                new WorkbookReadOperation.GetTables("tables", new TableSelection.All()),
                new WorkbookReadOperation.AnalyzeAutofilterHealth(
                    "autofilter-health", new SheetSelection.All()),
                new WorkbookReadOperation.AnalyzeTableHealth(
                    "table-health", new TableSelection.All())));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                workbookPath.toString(), workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadResult.AutofiltersResult(
                    "autofilters",
                    "Budget",
                    List.of(new AutofilterEntryReport.SheetOwned("E1:F3"))),
                new WorkbookReadResult.TablesResult(
                    "tables",
                    List.of(
                        new TableEntryReport(
                            "BudgetTable",
                            "Budget",
                            "A1:C4",
                            1,
                            0,
                            List.of("Item", "Amount", "Billable"),
                            new TableStyleReport.None(),
                            true))),
                new WorkbookReadResult.AutofilterHealthResult(
                    "autofilter-health",
                    new AutofilterHealthReport(
                        1, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of())),
                new WorkbookReadResult.TableHealthResult(
                    "table-health",
                    new TableHealthReport(
                        1, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of()))));

    assertDoesNotThrow(
        () -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }

  @Test
  void acceptsSuccessResponsesWithPivotReads(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("pivot.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs("pivot.xlsx", workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadResult.PivotTablesResult("pivots", List.of(pivotReport())),
                new WorkbookReadResult.PivotTableHealthResult(
                    "pivot-health",
                    new PivotTableHealthReport(
                        1, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of()))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsWorkflowOutcomeShapeForPivotReads(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("pivot.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadOperation.GetPivotTables(
                    "pivots", new PivotTableSelection.ByNames(List.of("OpsPivot"))),
                new WorkbookReadOperation.AnalyzePivotTableHealth(
                    "pivot-health", new PivotTableSelection.All())));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                workbookPath.toString(), workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadResult.PivotTablesResult("pivots", List.of(pivotReport())),
                new WorkbookReadResult.PivotTableHealthResult(
                    "pivot-health",
                    new PivotTableHealthReport(
                        1, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of()))));

    assertDoesNotThrow(
        () -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }

  @Test
  void acceptsSuccessResponsesWithAdvancedPhaseTwoReadShapes(@TempDir Path tempDirectory)
      throws IOException {
    Path workbookPath = tempDirectory.resolve("advanced.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindResponse.CommentReport anchoredComment =
        new GridGrindResponse.CommentReport(
            "Review",
            "GridGrind",
            true,
            List.of(
                new RichTextRunReport(
                    "Review",
                    new CellFontReport(
                        false,
                        false,
                        "Calibri",
                        new FontHeightReport(220, new BigDecimal("11")),
                        rgb("#C00000"),
                        false,
                        false))),
            new CommentAnchorReport(1, 2, 4, 6));
    GridGrindResponse.CellStyleReport advancedStyle =
        new GridGrindResponse.CellStyleReport(
            "General",
            new CellAlignmentReport(
                false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
            new CellFontReport(
                false,
                false,
                "Calibri",
                new FontHeightReport(220, new BigDecimal("11")),
                indexed(8),
                false,
                false),
            new CellFillReport(
                ExcelFillPattern.NONE,
                null,
                null,
                new CellGradientFillReport(
                    "LINEAR",
                    45.0d,
                    null,
                    null,
                    null,
                    null,
                    List.of(
                        new CellGradientStopReport(0.0d, rgb("#1F497D")),
                        new CellGradientStopReport(1.0d, themed(4, 0.45d))))),
            new CellBorderReport(
                new CellBorderSideReport(ExcelBorderStyle.NONE, indexed(16)),
                new CellBorderSideReport(ExcelBorderStyle.NONE, null),
                new CellBorderSideReport(ExcelBorderStyle.NONE, null),
                new CellBorderSideReport(ExcelBorderStyle.NONE, null)),
            new CellProtectionReport(true, false));

    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                "advanced.xlsx", workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadResult.WorkbookProtectionResult(
                    "workbook-protection",
                    new WorkbookProtectionReport(true, false, true, true, false)),
                new WorkbookReadResult.CellsResult(
                    "cells",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellReport.TextReport(
                            "A1",
                            "STRING",
                            "Review",
                            advancedStyle,
                            null,
                            anchoredComment,
                            "Review",
                            List.of(
                                new RichTextRunReport(
                                    "Review",
                                    new CellFontReport(
                                        false,
                                        false,
                                        "Calibri",
                                        new FontHeightReport(220, new BigDecimal("11")),
                                        rgb("#C00000"),
                                        false,
                                        false)))))),
                new WorkbookReadResult.CommentsResult(
                    "comments",
                    "Budget",
                    List.of(new GridGrindResponse.CellCommentReport("A1", anchoredComment))),
                new WorkbookReadResult.PrintLayoutResult(
                    "print-layout",
                    new PrintLayoutReport(
                        "Budget",
                        new PrintAreaReport.Range("A1:C20"),
                        ExcelPrintOrientation.LANDSCAPE,
                        new PrintScalingReport.Fit(1, 0),
                        new PrintTitleRowsReport.Band(0, 0),
                        new PrintTitleColumnsReport.None(),
                        new HeaderFooterTextReport("Budget", "Quarterly", ""),
                        new HeaderFooterTextReport("", "Confidential", "Page &P"),
                        new PrintSetupReport(
                            new PrintMarginsReport(0.5d, 0.5d, 0.75d, 0.75d, 0.3d, 0.3d),
                            false,
                            true,
                            false,
                            9,
                            false,
                            true,
                            2,
                            true,
                            3,
                            List.of(4, 9),
                            List.of(1)))),
                new WorkbookReadResult.AutofiltersResult(
                    "autofilters",
                    "Budget",
                    List.of(
                        new AutofilterEntryReport.TableOwned(
                            "A1:C5",
                            "BudgetTable",
                            List.of(
                                new AutofilterFilterColumnReport(
                                    1L,
                                    true,
                                    new AutofilterFilterCriterionReport.Values(
                                        List.of("Open", "Closed"), true))),
                            new AutofilterSortStateReport(
                                "A2:C5",
                                false,
                                false,
                                "",
                                List.of(
                                    new AutofilterSortConditionReport(
                                        "B2:B5", true, "", null, null)))))),
                new WorkbookReadResult.TablesResult(
                    "tables",
                    List.of(
                        new TableEntryReport(
                            "BudgetTable",
                            "Budget",
                            "A1:C5",
                            1,
                            1,
                            List.of("Item", "Status", "Owner"),
                            List.of(
                                new TableColumnReport(1L, "Item", "", "", "", ""),
                                new TableColumnReport(2L, "Status", "", "", "sum", ""),
                                new TableColumnReport(
                                    3L,
                                    "Owner",
                                    "owner_unique",
                                    "",
                                    "",
                                    "CONCAT([@Owner],\"-\",[@Status])")),
                            new TableStyleReport.Named(
                                "TableStyleMedium2", false, false, true, false),
                            true,
                            "Team queue",
                            true,
                            false,
                            true,
                            "HeaderStyle",
                            "DataStyle",
                            "TotalsStyle")))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsSuccessResponsesWithDrawingReadShapes(@TempDir Path tempDirectory)
      throws IOException {
    Path workbookPath = tempDirectory.resolve("drawing.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                "drawing.xlsx", workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadResult.DrawingObjectsResult(
                    "drawing-objects",
                    "Ops",
                    List.of(
                        new DrawingObjectReport.Picture(
                            "OpsPicture",
                            twoCellAnchor(),
                            ExcelPictureFormat.PNG,
                            "image/png",
                            68L,
                            "abc123",
                            1,
                            1,
                            "Queue preview"),
                        new DrawingObjectReport.Shape(
                            "OpsShape",
                            new DrawingAnchorReport.OneCell(
                                new DrawingMarkerReport(2, 3, 0, 0),
                                914_400L,
                                457_200L,
                                ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE),
                            ExcelDrawingShapeKind.SIMPLE_SHAPE,
                            "rect",
                            "Queue",
                            0),
                        new DrawingObjectReport.EmbeddedObject(
                            "OpsEmbed",
                            new DrawingAnchorReport.Absolute(
                                0L,
                                0L,
                                914_400L,
                                914_400L,
                                ExcelDrawingAnchorBehavior.DONT_MOVE_AND_RESIZE),
                            ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                            "Ops payload",
                            "ops-payload.txt",
                            "open",
                            "application/octet-stream",
                            17L,
                            "def456",
                            ExcelPictureFormat.PNG,
                            68L,
                            "preview789"))),
                new WorkbookReadResult.DrawingObjectPayloadResult(
                    "drawing-payload",
                    "Ops",
                    new DrawingObjectPayloadReport.EmbeddedObject(
                        "OpsEmbed",
                        ExcelEmbeddedObjectPackagingKind.OLE10_NATIVE,
                        "application/octet-stream",
                        "ops-payload.txt",
                        "def456",
                        "R3JpZEdyaW5kIHBheWxvYWQ=",
                        "Ops payload",
                        "open"))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsWorkflowOutcomeShapeForDrawingReads(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("drawing.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadOperation.GetDrawingObjects("drawing-objects", "Ops"),
                new WorkbookReadOperation.GetDrawingObjectPayload(
                    "drawing-payload", "Ops", "OpsPicture")));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                workbookPath.toString(), workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadResult.DrawingObjectsResult(
                    "drawing-objects",
                    "Ops",
                    List.of(
                        new DrawingObjectReport.Picture(
                            "OpsPicture",
                            twoCellAnchor(),
                            ExcelPictureFormat.PNG,
                            "image/png",
                            68L,
                            "abc123",
                            1,
                            1,
                            null))),
                new WorkbookReadResult.DrawingObjectPayloadResult(
                    "drawing-payload",
                    "Ops",
                    new DrawingObjectPayloadReport.Picture(
                        "OpsPicture",
                        ExcelPictureFormat.PNG,
                        "image/png",
                        "OpsPicture.png",
                        "abc123",
                        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=",
                        null))));

    assertDoesNotThrow(
        () -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }

  @Test
  void acceptsSuccessResponsesWithChartReadShapes(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("chart.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs("chart.xlsx", workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadResult.DrawingObjectsResult(
                    "drawing-objects",
                    "Ops",
                    List.of(
                        new DrawingObjectReport.Chart(
                            "OpsChart", twoCellAnchor(), true, List.of("BAR"), "Roadmap"))),
                new WorkbookReadResult.ChartsResult("charts", "Ops", List.of(chartReport()))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsWorkflowOutcomeShapeForChartReads(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("chart.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadOperation.GetDrawingObjects("drawing-objects", "Ops"),
                new WorkbookReadOperation.GetCharts("charts", "Ops")));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                workbookPath.toString(), workbookPath.toString()),
            List.of(),
            List.of(
                new WorkbookReadResult.DrawingObjectsResult(
                    "drawing-objects",
                    "Ops",
                    List.of(
                        new DrawingObjectReport.Chart(
                            "OpsChart", twoCellAnchor(), true, List.of("BAR"), "Roadmap"))),
                new WorkbookReadResult.ChartsResult("charts", "Ops", List.of(chartReport()))));

    assertDoesNotThrow(
        () -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }

  @Test
  void acceptsWorkbookShapeWithCharts() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops").setCell("A1", ExcelCellValue.text("Month"));
      workbook.getOrCreateSheet("Ops").setCell("B1", ExcelCellValue.text("Actual"));
      workbook.getOrCreateSheet("Ops").setCell("A2", ExcelCellValue.text("Jan"));
      workbook.getOrCreateSheet("Ops").setCell("B2", ExcelCellValue.number(12.0d));
      workbook.getOrCreateSheet("Ops").setCell("A3", ExcelCellValue.text("Feb"));
      workbook.getOrCreateSheet("Ops").setCell("B3", ExcelCellValue.number(18.0d));
      workbook.getOrCreateSheet("Ops").setCell("A4", ExcelCellValue.text("Mar"));
      workbook.getOrCreateSheet("Ops").setCell("B4", ExcelCellValue.number(15.0d));
      new WorkbookCommandExecutor()
          .apply(
              workbook,
              List.of(
                  new WorkbookCommand.SetChart(
                      "Ops",
                      new ExcelChartDefinition.Bar(
                          "OpsChart",
                          new ExcelDrawingAnchor.TwoCell(
                              new ExcelDrawingMarker(0, 0, 0, 0),
                              new ExcelDrawingMarker(2, 8, 0, 0),
                              null),
                          new ExcelChartDefinition.Title.Text("Roadmap"),
                          new ExcelChartDefinition.Legend.Visible(
                              ExcelChartLegendPosition.TOP_RIGHT),
                          ExcelChartDisplayBlanksAs.SPAN,
                          false,
                          true,
                          ExcelChartBarDirection.COLUMN,
                          List.of(
                              new ExcelChartDefinition.Series(
                                  new ExcelChartDefinition.Title.Text("Actual"),
                                  new ExcelChartDefinition.DataSource("Ops!$A$2:$A$4"),
                                  new ExcelChartDefinition.DataSource("Ops!$B$2:$B$4")))))));

      assertDoesNotThrow(() -> WorkbookInvariantChecks.requireWorkbookShape(workbook));
    }
  }

  @Test
  void acceptsWorkbookShapeWithPivots() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.text("Month"));
      workbook.getOrCreateSheet("Budget").setCell("B1", ExcelCellValue.text("Actual"));
      workbook.getOrCreateSheet("Budget").setCell("A2", ExcelCellValue.text("Jan"));
      workbook.getOrCreateSheet("Budget").setCell("B2", ExcelCellValue.number(12.0d));
      workbook.getOrCreateSheet("Budget").setCell("A3", ExcelCellValue.text("Feb"));
      workbook.getOrCreateSheet("Budget").setCell("B3", ExcelCellValue.number(18.0d));
      workbook.getOrCreateSheet("Budget").setCell("A4", ExcelCellValue.text("Mar"));
      workbook.getOrCreateSheet("Budget").setCell("B4", ExcelCellValue.number(15.0d));
      workbook.getOrCreateSheet("Pivot");
      new WorkbookCommandExecutor()
          .apply(
              workbook,
              List.of(
                  new WorkbookCommand.SetPivotTable(
                      new ExcelPivotTableDefinition(
                          "OpsPivot",
                          "Pivot",
                          new ExcelPivotTableDefinition.Source.Range("Budget", "A1:B4"),
                          new ExcelPivotTableDefinition.Anchor("C5"),
                          List.of("Month"),
                          List.of(),
                          List.of(),
                          List.of(
                              new ExcelPivotTableDefinition.DataField(
                                  "Actual",
                                  ExcelPivotDataConsolidateFunction.SUM,
                                  "Total Actual",
                                  null))))));

      assertDoesNotThrow(() -> WorkbookInvariantChecks.requireWorkbookShape(workbook));
    }
  }

  @Test
  void acceptsWorkflowOutcomeShapeWithPackageSecurityRead(@TempDir Path tempDirectory)
      throws IOException {
    Path workbookPath = tempDirectory.resolve("secured.xlsx");
    Files.writeString(workbookPath, "seed");

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.ExistingFile(
                workbookPath.toString(),
                new dev.erst.gridgrind.protocol.dto.OoxmlOpenSecurityInput("GridGrind-2026")),
            new GridGrindRequest.WorkbookPersistence.None(),
            List.of(),
            List.of(new WorkbookReadOperation.GetPackageSecurity("security")));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.NotSaved(),
            List.of(),
            List.of(
                new WorkbookReadResult.PackageSecurityResult(
                    "security",
                    new dev.erst.gridgrind.protocol.dto.OoxmlPackageSecurityReport(
                        new dev.erst.gridgrind.protocol.dto.OoxmlEncryptionReport(
                            true,
                            dev.erst.gridgrind.excel.ExcelOoxmlEncryptionMode.AGILE,
                            "aes",
                            "sha512",
                            "ChainingModeCBC",
                            256,
                            16,
                            100000),
                        List.of(
                            new dev.erst.gridgrind.protocol.dto.OoxmlSignatureReport(
                                "/_xmlsignatures/sig1.xml",
                                "CN=GridGrind Signing",
                                "CN=GridGrind Signing",
                                "01",
                                dev.erst.gridgrind.excel.ExcelOoxmlSignatureState.VALID))))));

    assertDoesNotThrow(
        () -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }

  private static ChartReport.Bar chartReport() {
    return new ChartReport.Bar(
        "OpsChart",
        twoCellAnchor(),
        new ChartReport.Title.Text("Roadmap"),
        new ChartReport.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        true,
        ExcelChartBarDirection.COLUMN,
        List.of(
            new ChartReport.Axis(
                dev.erst.gridgrind.excel.ExcelChartAxisKind.CATEGORY,
                dev.erst.gridgrind.excel.ExcelChartAxisPosition.BOTTOM,
                dev.erst.gridgrind.excel.ExcelChartAxisCrosses.AUTO_ZERO,
                true),
            new ChartReport.Axis(
                dev.erst.gridgrind.excel.ExcelChartAxisKind.VALUE,
                dev.erst.gridgrind.excel.ExcelChartAxisPosition.LEFT,
                dev.erst.gridgrind.excel.ExcelChartAxisCrosses.AUTO_ZERO,
                true)),
        List.of(
            new ChartReport.Series(
                new ChartReport.Title.Text("Actual"),
                new ChartReport.DataSource.StringReference(
                    "Ops!$A$2:$A$4", List.of("Jan", "Feb", "Mar")),
                new ChartReport.DataSource.NumericReference(
                    "Ops!$B$2:$B$4", "General", List.of("12", "18", "15")))));
  }

  private static PivotTableReport.Supported pivotReport() {
    return new PivotTableReport.Supported(
        "OpsPivot",
        "Budget",
        new PivotTableReport.Anchor("F4", "F4:H8"),
        new PivotTableReport.Source.Range("Budget", "A1:C4"),
        List.of(new PivotTableReport.Field(0, "Month")),
        List.of(),
        List.of(),
        List.of(
            new PivotTableReport.DataField(
                2, "Actual", ExcelPivotDataConsolidateFunction.SUM, "Total Actual", "General")),
        false);
  }

  private static GridGrindResponse.CellStyleReport defaultStyle() {
    return new GridGrindResponse.CellStyleReport(
        "General",
        new CellAlignmentReport(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new CellFontReport(
            false,
            false,
            "Calibri",
            new FontHeightReport(220, new BigDecimal("11")),
            null,
            false,
            false),
        new CellFillReport(ExcelFillPattern.NONE, null, null),
        new CellBorderReport(
            new CellBorderSideReport(ExcelBorderStyle.NONE, null),
            new CellBorderSideReport(ExcelBorderStyle.NONE, null),
            new CellBorderSideReport(ExcelBorderStyle.NONE, null),
            new CellBorderSideReport(ExcelBorderStyle.NONE, null)),
        new CellProtectionReport(true, false));
  }

  private static GridGrindResponse.CellReport.TextReport textCell(String address, String value) {
    return new GridGrindResponse.CellReport.TextReport(
        address, "STRING", value, defaultStyle(), null, null, value, null);
  }

  private static CellColorReport rgb(String rgb) {
    return new CellColorReport(rgb);
  }

  private static CellColorReport indexed(int indexed) {
    return new CellColorReport(null, null, indexed, null);
  }

  private static CellColorReport themed(int theme, double tint) {
    return new CellColorReport(null, theme, null, tint);
  }

  private static DrawingAnchorReport.TwoCell twoCellAnchor() {
    return new DrawingAnchorReport.TwoCell(
        new DrawingMarkerReport(0, 0, 0, 0),
        new DrawingMarkerReport(2, 3, 0, 0),
        ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
  }

  private static SheetProtectionSettings protocolProtectionSettings() {
    return new SheetProtectionSettings(
        false, true, false, true, false, true, false, true, false, true, false, true, false, true,
        false);
  }

  private static ExcelSheetProtectionSettings protectionSettings() {
    return new ExcelSheetProtectionSettings(
        false, true, false, true, false, true, false, true, false, true, false, true, false, true,
        false);
  }
}
