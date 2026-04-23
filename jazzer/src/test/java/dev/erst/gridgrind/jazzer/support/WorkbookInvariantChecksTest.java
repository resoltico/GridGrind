package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.assertThat;
import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.inspect;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.AnalysisFindingCode;
import dev.erst.gridgrind.contract.dto.AnalysisSeverity;
import dev.erst.gridgrind.contract.dto.ArrayFormulaReport;
import dev.erst.gridgrind.contract.dto.AutofilterEntryReport;
import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnReport;
import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.contract.dto.AutofilterHealthReport;
import dev.erst.gridgrind.contract.dto.AutofilterSortConditionReport;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateReport;
import dev.erst.gridgrind.contract.dto.CellAlignmentReport;
import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellColorReport;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellGradientFillReport;
import dev.erst.gridgrind.contract.dto.CellGradientStopReport;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.CommentAnchorReport;
import dev.erst.gridgrind.contract.dto.CustomXmlDataBindingReport;
import dev.erst.gridgrind.contract.dto.CustomXmlExportReport;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.contract.dto.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingLocator;
import dev.erst.gridgrind.contract.dto.CustomXmlMappingReport;
import dev.erst.gridgrind.contract.dto.DataValidationEntryReport;
import dev.erst.gridgrind.contract.dto.DataValidationHealthReport;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorReport;
import dev.erst.gridgrind.contract.dto.DrawingMarkerReport;
import dev.erst.gridgrind.contract.dto.DrawingObjectPayloadReport;
import dev.erst.gridgrind.contract.dto.DrawingObjectReport;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextReport;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.PaneReport;
import dev.erst.gridgrind.contract.dto.PivotTableHealthReport;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.PrintAreaReport;
import dev.erst.gridgrind.contract.dto.PrintLayoutReport;
import dev.erst.gridgrind.contract.dto.PrintMarginsReport;
import dev.erst.gridgrind.contract.dto.PrintScalingReport;
import dev.erst.gridgrind.contract.dto.PrintSetupReport;
import dev.erst.gridgrind.contract.dto.PrintTitleColumnsReport;
import dev.erst.gridgrind.contract.dto.PrintTitleRowsReport;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.RichTextRunReport;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.TableColumnReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.TableHealthReport;
import dev.erst.gridgrind.contract.dto.TableStyleReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.ExcelPivotTableDefinition;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
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
                    1,
                    "step-01-set-cell",
                    "SET_CELL",
                    "Formula references same-request sheet names with spaces.")),
            List.of(new AssertionResult("assert-total", "EXPECT_PRESENT")),
            List.of(
                new InspectionResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of("Budget"), 1, false)),
                new InspectionResult.NamedRangesResult(
                    "ranges",
                    List.of(
                        new GridGrindResponse.NamedRangeReport.RangeReport(
                            "BudgetTotal",
                            new NamedRangeScope.Workbook(),
                            "Budget!$B$4",
                            new NamedRangeTarget("Budget", "B4")))),
                new InspectionResult.CellsResult(
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
                new InspectionResult.WindowResult(
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
                new InspectionResult.MergedRegionsResult(
                    "merged", "Budget", List.of(new GridGrindResponse.MergedRegionReport("A1:B1"))),
                new InspectionResult.HyperlinksResult(
                    "hyperlinks",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellHyperlinkReport(
                            "A1", new HyperlinkTarget.Url("https://example.com/report")))),
                new InspectionResult.CommentsResult(
                    "comments",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellCommentReport(
                            "A1",
                            new GridGrindResponse.CommentReport("Review", "GridGrind", true)))),
                new InspectionResult.SheetLayoutResult(
                    "layout",
                    new GridGrindResponse.SheetLayoutReport(
                        "Budget",
                        new PaneReport.Frozen(1, 1, 1, 1),
                        125,
                        dev.erst.gridgrind.contract.dto.SheetPresentationReport.defaults(),
                        List.of(new GridGrindResponse.ColumnLayoutReport(0, 12.5, false, 0, false)),
                        List.of(new GridGrindResponse.RowLayoutReport(0, 18.0, false, 0, false)))),
                new InspectionResult.PrintLayoutResult(
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
                new InspectionResult.DataValidationsResult(
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
                new InspectionResult.FormulaSurfaceResult(
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
                new InspectionResult.SheetSchemaResult(
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
                new InspectionResult.NamedRangeSurfaceResult(
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
                new InspectionResult.FormulaHealthResult(
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
                new InspectionResult.DataValidationHealthResult(
                    "data-validation-health",
                    new DataValidationHealthReport(
                        1, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of()))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsWorkflowOutcomeShapeForOrderedReads(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");

    WorkbookPlan request =
        saveAsRequest(
            workbookPath,
            inspect(
                "summary",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetWorkbookSummary()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Budget", List.of("A1")),
                new InspectionQuery.GetCells()),
            inspect(
                "data-validations",
                new RangeSelector.AllOnSheet("Budget"),
                new InspectionQuery.GetDataValidations()),
            inspect(
                "hyperlinks",
                new CellSelector.AllUsedInSheet("Budget"),
                new InspectionQuery.GetHyperlinks()),
            inspect(
                "data-validation-health",
                new SheetSelector.All(),
                new InspectionQuery.AnalyzeDataValidationHealth()),
            inspect(
                "named-range-health",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                new InspectionQuery.AnalyzeNamedRangeHealth()));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                workbookPath.toString(), workbookPath.toString()),
            List.of(),
            List.of(),
            List.of(
                new InspectionResult.WorkbookSummaryResult(
                    "summary",
                    new GridGrindResponse.WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of("Budget"), 1, false)),
                new InspectionResult.CellsResult(
                    "cells", "Budget", List.of(textCell("A1", "Report"))),
                new InspectionResult.DataValidationsResult(
                    "data-validations",
                    "Budget",
                    List.of(
                        new DataValidationEntryReport.Unsupported(
                            List.of("A2:A5"), "MISSING_FORMULA", "Formula is missing"))),
                new InspectionResult.HyperlinksResult(
                    "hyperlinks",
                    "Budget",
                    List.of(
                        new GridGrindResponse.CellHyperlinkReport(
                            "A1", new HyperlinkTarget.Url("https://example.com/report")))),
                new InspectionResult.DataValidationHealthResult(
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
                new InspectionResult.NamedRangeHealthResult(
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
            List.of(),
            List.of(
                new InspectionResult.HyperlinksResult(
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
            List.of(),
            List.of(
                new InspectionResult.SheetSummaryResult(
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
            List.of(),
            List.of(
                new InspectionResult.AutofiltersResult(
                    "autofilters",
                    "Budget",
                    List.of(
                        new AutofilterEntryReport.SheetOwned("E1:F3"),
                        new AutofilterEntryReport.TableOwned("A1:C4", "BudgetTable"))),
                new InspectionResult.TablesResult(
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
                new InspectionResult.AutofilterHealthResult(
                    "autofilter-health",
                    new AutofilterHealthReport(
                        2, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of())),
                new InspectionResult.TableHealthResult(
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

    WorkbookPlan request =
        saveAsRequest(
            workbookPath,
            inspect(
                "autofilters",
                new SheetSelector.ByName("Budget"),
                new InspectionQuery.GetAutofilters()),
            inspect("tables", new TableSelector.All(), new InspectionQuery.GetTables()),
            inspect(
                "autofilter-health",
                new SheetSelector.All(),
                new InspectionQuery.AnalyzeAutofilterHealth()),
            inspect(
                "table-health", new TableSelector.All(), new InspectionQuery.AnalyzeTableHealth()));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                workbookPath.toString(), workbookPath.toString()),
            List.of(),
            List.of(),
            List.of(
                new InspectionResult.AutofiltersResult(
                    "autofilters",
                    "Budget",
                    List.of(new AutofilterEntryReport.SheetOwned("E1:F3"))),
                new InspectionResult.TablesResult(
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
                new InspectionResult.AutofilterHealthResult(
                    "autofilter-health",
                    new AutofilterHealthReport(
                        1, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of())),
                new InspectionResult.TableHealthResult(
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
            List.of(),
            List.of(
                new InspectionResult.PivotTablesResult("pivots", List.of(pivotReport())),
                new InspectionResult.PivotTableHealthResult(
                    "pivot-health",
                    new PivotTableHealthReport(
                        1, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of()))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsWorkflowOutcomeShapeForPivotReads(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("pivot.xlsx");
    Files.writeString(workbookPath, "seed");

    WorkbookPlan request =
        saveAsRequest(
            workbookPath,
            inspect(
                "pivots",
                new PivotTableSelector.ByNames(List.of("OpsPivot")),
                new InspectionQuery.GetPivotTables()),
            inspect(
                "pivot-health",
                new PivotTableSelector.All(),
                new InspectionQuery.AnalyzePivotTableHealth()));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                workbookPath.toString(), workbookPath.toString()),
            List.of(),
            List.of(),
            List.of(
                new InspectionResult.PivotTablesResult("pivots", List.of(pivotReport())),
                new InspectionResult.PivotTableHealthResult(
                    "pivot-health",
                    new PivotTableHealthReport(
                        1, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of()))));

    assertDoesNotThrow(
        () -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }

  @Test
  void acceptsWorkflowOutcomeShapeForAssertions(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("assertions.xlsx");
    Files.writeString(workbookPath, "seed");

    WorkbookPlan request =
        ProtocolStepSupport.request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
            List.of(),
            List.of(
                assertThat(
                    "assert-total",
                    new SheetSelector.ByName("Budget"),
                    new Assertion.AnalysisMaxSeverity(
                        new InspectionQuery.AnalyzeFormulaHealth(), AnalysisSeverity.ERROR))),
            List.of(
                inspect(
                    "sheet",
                    new SheetSelector.ByName("Budget"),
                    new InspectionQuery.GetSheetSummary())));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                workbookPath.toString(), workbookPath.toString()),
            List.of(),
            List.of(new AssertionResult("assert-total", "EXPECT_ANALYSIS_MAX_SEVERITY")),
            List.of(
                new InspectionResult.SheetSummaryResult(
                    "sheet",
                    new GridGrindResponse.SheetSummaryReport(
                        "Budget",
                        ExcelSheetVisibility.VISIBLE,
                        new GridGrindResponse.SheetProtectionReport.Unprotected(),
                        0,
                        -1,
                        -1))));

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
            List.of(),
            List.of(
                new InspectionResult.WorkbookProtectionResult(
                    "workbook-protection",
                    new WorkbookProtectionReport(true, false, true, true, false)),
                new InspectionResult.CellsResult(
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
                new InspectionResult.CommentsResult(
                    "comments",
                    "Budget",
                    List.of(new GridGrindResponse.CellCommentReport("A1", anchoredComment))),
                new InspectionResult.PrintLayoutResult(
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
                new InspectionResult.AutofiltersResult(
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
                new InspectionResult.TablesResult(
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
            List.of(),
            List.of(
                new InspectionResult.DrawingObjectsResult(
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
                new InspectionResult.DrawingObjectPayloadResult(
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

    WorkbookPlan request =
        saveAsRequest(
            workbookPath,
            inspect(
                "drawing-objects",
                new DrawingObjectSelector.AllOnSheet("Ops"),
                new InspectionQuery.GetDrawingObjects()),
            inspect(
                "drawing-payload",
                new DrawingObjectSelector.ByName("Ops", "OpsPicture"),
                new InspectionQuery.GetDrawingObjectPayload()));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                workbookPath.toString(), workbookPath.toString()),
            List.of(),
            List.of(),
            List.of(
                new InspectionResult.DrawingObjectsResult(
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
                new InspectionResult.DrawingObjectPayloadResult(
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
            List.of(),
            List.of(
                new InspectionResult.DrawingObjectsResult(
                    "drawing-objects",
                    "Ops",
                    List.of(
                        new DrawingObjectReport.Chart(
                            "OpsChart", twoCellAnchor(), true, List.of("BAR"), "Roadmap"))),
                new InspectionResult.ChartsResult("charts", "Ops", List.of(chartReport()))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsWorkflowOutcomeShapeForChartReads(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("chart.xlsx");
    Files.writeString(workbookPath, "seed");

    WorkbookPlan request =
        saveAsRequest(
            workbookPath,
            inspect(
                "drawing-objects",
                new DrawingObjectSelector.AllOnSheet("Ops"),
                new InspectionQuery.GetDrawingObjects()),
            inspect(
                "charts", new ChartSelector.AllOnSheet("Ops"), new InspectionQuery.GetCharts()));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.SavedAs(
                workbookPath.toString(), workbookPath.toString()),
            List.of(),
            List.of(),
            List.of(
                new InspectionResult.DrawingObjectsResult(
                    "drawing-objects",
                    "Ops",
                    List.of(
                        new DrawingObjectReport.Chart(
                            "OpsChart", twoCellAnchor(), true, List.of("BAR"), "Roadmap"))),
                new InspectionResult.ChartsResult("charts", "Ops", List.of(chartReport()))));

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
                      new ExcelChartDefinition(
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
                          List.of(
                              new ExcelChartDefinition.Bar(
                                  true,
                                  ExcelChartBarDirection.COLUMN,
                                  ExcelChartBarGrouping.CLUSTERED,
                                  null,
                                  null,
                                  excelCategoryAxes(),
                                  List.of(
                                      new ExcelChartDefinition.Series(
                                          new ExcelChartDefinition.Title.Text("Actual"),
                                          new ExcelChartDefinition.DataSource.Reference(
                                              "Ops!$A$2:$A$4"),
                                          new ExcelChartDefinition.DataSource.Reference(
                                              "Ops!$B$2:$B$4"),
                                          null,
                                          null,
                                          null,
                                          null))))))));

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

    WorkbookPlan request =
        existingRequest(
            new WorkbookPlan.WorkbookSource.ExistingFile(
                workbookPath.toString(),
                new dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput("GridGrind-2026")),
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetPackageSecurity()));
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.NotSaved(),
            List.of(),
            List.of(),
            List.of(
                new InspectionResult.PackageSecurityResult(
                    "security",
                    new dev.erst.gridgrind.contract.dto.OoxmlPackageSecurityReport(
                        new dev.erst.gridgrind.contract.dto.OoxmlEncryptionReport(
                            true,
                            dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode.AGILE,
                            "aes",
                            "sha512",
                            "ChainingModeCBC",
                            256,
                            16,
                            100000),
                        List.of(
                            new dev.erst.gridgrind.contract.dto.OoxmlSignatureReport(
                                "/_xmlsignatures/sig1.xml",
                                "CN=GridGrind Signing",
                                "CN=GridGrind Signing",
                                "01",
                                dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState
                                    .VALID))))));

    assertDoesNotThrow(
        () -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }

  @Test
  void acceptsWorkflowOutcomeShapeForCustomXmlArrayFormulaAndSignatureLineReads(
      @TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("advanced.xlsx");
    Files.writeString(workbookPath, "seed");

    WorkbookPlan request =
        existingRequest(
            new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
            inspect(
                "custom-xml-mappings",
                new WorkbookSelector.Current(),
                new InspectionQuery.GetCustomXmlMappings()),
            inspect(
                "custom-xml-export",
                new WorkbookSelector.Current(),
                new InspectionQuery.ExportCustomXmlMapping(
                    new CustomXmlMappingLocator(1L, "BudgetMap"), true, "UTF-8")),
            inspect(
                "array-formulas",
                new SheetSelector.ByName("Ops"),
                new InspectionQuery.GetArrayFormulas()),
            inspect(
                "drawing-objects",
                new DrawingObjectSelector.AllOnSheet("Ops"),
                new InspectionQuery.GetDrawingObjects()));
    CustomXmlMappingReport mapping = customXmlMappingReport();
    GridGrindResponse.Success response =
        new GridGrindResponse.Success(
            GridGrindProtocolVersion.V1,
            new GridGrindResponse.PersistenceOutcome.NotSaved(),
            List.of(),
            List.of(),
            List.of(
                new InspectionResult.CustomXmlMappingsResult(
                    "custom-xml-mappings", List.of(mapping)),
                new InspectionResult.CustomXmlExportResult(
                    "custom-xml-export",
                    new CustomXmlExportReport(
                        mapping,
                        "UTF-8",
                        true,
                        "<BudgetMap><Owner>Ada Lovelace</Owner></BudgetMap>")),
                new InspectionResult.ArrayFormulasResult(
                    "array-formulas",
                    List.of(new ArrayFormulaReport("Ops", "D2:D4", "D2", "B2:B4*C2:C4", false))),
                new InspectionResult.DrawingObjectsResult(
                    "drawing-objects", "Ops", List.of(signatureLineDrawingObjectReport()))));

    assertDoesNotThrow(
        () -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }

  private static ChartReport chartReport() {
    return new ChartReport(
        "OpsChart",
        twoCellAnchor(),
        new ChartReport.Title.Text("Roadmap"),
        new ChartReport.Legend.Visible(ExcelChartLegendPosition.TOP_RIGHT),
        ExcelChartDisplayBlanksAs.SPAN,
        false,
        List.of(
            new ChartReport.Bar(
                true,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.CLUSTERED,
                null,
                null,
                chartCategoryAxes(),
                List.of(
                    new ChartReport.Series(
                        new ChartReport.Title.Text("Actual"),
                        new ChartReport.DataSource.StringReference(
                            "Ops!$A$2:$A$4", List.of("Jan", "Feb", "Mar")),
                        new ChartReport.DataSource.NumericReference(
                            "Ops!$B$2:$B$4", "General", List.of("12", "18", "15")),
                        null,
                        null,
                        null,
                        null)))));
  }

  private static CustomXmlMappingReport customXmlMappingReport() {
    return new CustomXmlMappingReport(
        1L,
        "BudgetMap",
        "BudgetMap",
        "schema-1",
        true,
        true,
        false,
        true,
        true,
        "urn:gridgrind:budget",
        "en-US",
        "budget-map.xsd",
        "<xs:schema/>",
        new CustomXmlDataBindingReport("BudgetBinding", false, 42L, "budget.xml", 1L),
        List.of(new CustomXmlLinkedCellReport("Ops", "A2", "/BudgetMap/Owner[1]", "string")),
        List.of(
            new CustomXmlLinkedTableReport(
                "Ops", "BudgetTable", "BudgetTable", "A1:B4", "/BudgetMap/Rows")));
  }

  private static DrawingObjectReport.SignatureLine signatureLineDrawingObjectReport() {
    return new DrawingObjectReport.SignatureLine(
        "BudgetSignature",
        twoCellAnchor(),
        "sig-setup-01",
        false,
        "Review the budget before signing.",
        "Ada Lovelace",
        "Finance",
        "ada@example.com",
        ExcelPictureFormat.PNG,
        "image/png",
        128L,
        "0123456789abcdef",
        320,
        120);
  }

  private static List<ChartReport.Axis> chartCategoryAxes() {
    return List.of(
        new ChartReport.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ChartReport.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  private static List<ExcelChartDefinition.Axis> excelCategoryAxes() {
    return List.of(
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.CATEGORY,
            ExcelChartAxisPosition.BOTTOM,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true),
        new ExcelChartDefinition.Axis(
            ExcelChartAxisKind.VALUE,
            ExcelChartAxisPosition.LEFT,
            ExcelChartAxisCrosses.AUTO_ZERO,
            true));
  }

  @SafeVarargs
  private static WorkbookPlan saveAsRequest(Path workbookPath, InspectionStep... inspections) {
    return ProtocolStepSupport.request(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
        List.of(),
        List.of(inspections));
  }

  @SafeVarargs
  private static WorkbookPlan existingRequest(
      WorkbookPlan.WorkbookSource.ExistingFile source, InspectionStep... inspections) {
    return ProtocolStepSupport.request(
        source, new WorkbookPlan.WorkbookPersistence.None(), List.of(), List.of(inspections));
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
