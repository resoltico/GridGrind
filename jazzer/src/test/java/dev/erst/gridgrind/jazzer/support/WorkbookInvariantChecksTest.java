package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.assertThat;
import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.inspect;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.*;
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
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.contract.dto.WorkbookResults;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.query.SheetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookAnalysisResult;
import dev.erst.gridgrind.contract.query.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookSurfaceInspectionResult;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingMarker;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
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
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlHashAlgorithm;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableDefinition;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Tests for WorkbookInvariantChecks success-shape validation. */
class WorkbookInvariantChecksTest {
  @Test
  void acceptsSuccessResponsesWithOrderedReads(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");
    CellStyleReport style = defaultStyle();

    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            writtenSaveAs("result.xlsx", workbookPath),
            List.of(
                new RequestWarning(
                    GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
                    1,
                    "step-01-set-cell",
                    "SET_CELL",
                    "Formula references same-request sheet names with spaces.")),
            List.of(new AssertionResult.Passed("assert-total", "EXPECT_NAMED_RANGE_PRESENT")),
            List.of(
                new WorkbookInspectionResult.WorkbookSummaryResult(
                    "summary",
                    new WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of("Budget"), 1, false)),
                new WorkbookInspectionResult.NamedRangesResult(
                    "ranges",
                    List.of(
                        new NamedRangeReport.RangeReport(
                            "BudgetTotal",
                            new NamedRangeScope.Workbook(),
                            "Budget!$B$4",
                            NamedRangeTarget.range("Budget", "B4")))),
                new SheetInspectionResult.CellsResult(
                    "cells",
                    "Budget",
                    List.of(
                        new dev.erst.gridgrind.contract.dto.CellReport.TextReport(
                            "A1",
                            java.util.Optional.of("Report"),
                            java.util.Optional.of(style),
                            java.util.Optional.of(
                                new HyperlinkTarget.Url("https://example.com/report")),
                            java.util.Optional.of(new CommentReport("Review", "GridGrind", true)),
                            java.util.Optional.of("Report"),
                            java.util.Optional.of(
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
                                            false))))))),
                new SheetInspectionResult.WindowResult(
                    "window",
                    new WindowReport.Dense(
                        "Budget",
                        "A1",
                        new WindowDimensionsReport(1, 1),
                        List.of(
                            new WindowRowReport(
                                0,
                                List.of(
                                    new dev.erst.gridgrind.contract.dto.CellReport.TextReport(
                                        "A1",
                                        java.util.Optional.empty(),
                                        java.util.Optional.of(style),
                                        java.util.Optional.empty(),
                                        java.util.Optional.empty(),
                                        java.util.Optional.of("Report"),
                                        java.util.Optional.empty())))))),
                new SheetInspectionResult.MergedRegionsResult(
                    "merged", "Budget", List.of(new MergedRegionReport("A1:B1"))),
                new SheetInspectionResult.HyperlinksResult(
                    "hyperlinks",
                    "Budget",
                    List.of(
                        new CellHyperlinkReport(
                            "A1", new HyperlinkTarget.Url("https://example.com/report")))),
                new SheetInspectionResult.CommentsResult(
                    "comments",
                    "Budget",
                    List.of(
                        new CellCommentReport(
                            "A1", new CommentReport("Review", "GridGrind", true)))),
                new SheetInspectionResult.SheetLayoutResult(
                    "layout",
                    new SheetLayoutReport(
                        "Budget",
                        new PaneReport.Frozen(1, 1, 1, 1),
                        125,
                        dev.erst.gridgrind.contract.dto.SheetPresentationReport.defaults(),
                        List.of(new ColumnLayoutReport(0, 12.5, false, 0, false)),
                        List.of(new RowLayoutReport(0, 18.0, false, 0, false)))),
                new SheetInspectionResult.PrintLayoutResult(
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
                new SheetInspectionResult.DataValidationsResult(
                    "data-validations",
                    "Budget",
                    List.of(
                        new DataValidationEntryReport.Supported(
                            List.of("A2:A5"),
                            new DataValidationEntryReport.DataValidationDefinitionReport(
                                new DataValidationRuleInput.WholeNumber(
                                    ExcelComparisonOperator.GREATER_OR_EQUAL,
                                    "1",
                                    Optional.empty()),
                                true,
                                false,
                                Optional.empty(),
                                Optional.empty())))),
                new WorkbookSurfaceInspectionResult.FormulaSurfaceResult(
                    "formula-surface",
                    new FormulaSurfaceReport(
                        1,
                        List.of(
                            new SheetFormulaSurfaceReport(
                                "Budget",
                                1,
                                1,
                                List.of(
                                    new FormulaPatternReport("SUM(B2:B3)", 1, List.of("B4"))))))),
                new WorkbookSurfaceInspectionResult.SheetSchemaResult(
                    "schema",
                    new SheetSchemaReport(
                        "Budget",
                        "A1",
                        2,
                        1,
                        1,
                        List.of(
                            new SchemaColumnReport(
                                0,
                                "A1",
                                "Item",
                                1,
                                0,
                                List.of(new TypeCountReport("TEXT", 1)),
                                "TEXT")))),
                new WorkbookSurfaceInspectionResult.NamedRangeSurfaceResult(
                    "named-range-surface",
                    new NamedRangeSurfaceReport(
                        1,
                        0,
                        1,
                        0,
                        List.of(
                            new NamedRangeSurfaceEntryReport(
                                "BudgetTotal",
                                new NamedRangeScope.Workbook(),
                                "Budget!$B$4",
                                NamedRangeBackingKind.RANGE)))),
                new WorkbookAnalysisResult.FormulaHealthResult(
                    "formula-health",
                    new FormulaHealthReport(
                        1,
                        new AnalysisSummaryReport(1, 0, 0, 1),
                        List.of(
                            new AnalysisFindingReport(
                                AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                                AnalysisSeverity.INFO,
                                "Volatile formula",
                                "Formula uses NOW().",
                                new AnalysisLocationReport.Cell("Budget", "B4"),
                                List.of("NOW()"))))),
                new WorkbookAnalysisResult.DataValidationHealthResult(
                    "data-validation-health",
                    new DataValidationHealthReport(
                        1, new AnalysisSummaryReport(0, 0, 0, 0), List.of()))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsFailureResponsesWithFailureCapablePersistence(@TempDir Path tempDirectory)
      throws IOException {
    Path workbookPath = tempDirectory.resolve("failed-save.xlsx");
    WorkbookPlan request = saveAsRequest(workbookPath);
    WorkbookResult.Failure response =
        WorkbookResults.failure(
            GridGrindProtocolVersion.V2,
            notWrittenSaveAs(workbookPath.toString()),
            GridGrindProblemDetail.Problem.of(
                GridGrindProblemCode.IO_ERROR,
                "Could not write workbook to " + workbookPath + ": disk full",
                new ProblemContext.PersistWorkbook(
                    ProblemContextRequestSurfaces.RequestShape.known("NEW", "SAVE_AS"),
                    ProblemContextWorkbookSurfaces.PersistenceReference.saveAs(
                        workbookPath.toString()))));
    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
    assertDoesNotThrow(
        () -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
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
                new WorkbookIntrospectionQuery.GetWorkbookSummary()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Budget", List.of("A1")),
                new SheetIntrospectionQuery.GetCells()),
            inspect(
                "data-validations",
                new RangeSelector.AllOnSheet("Budget"),
                new SheetIntrospectionQuery.GetDataValidations()),
            inspect(
                "hyperlinks",
                new CellSelector.AllUsedInSheet("Budget"),
                new SheetIntrospectionQuery.GetHyperlinks()),
            inspect(
                "data-validation-health",
                new SheetSelector.All(),
                new InspectionAnalysisQuery.AnalyzeDataValidationHealth()),
            inspect(
                "named-range-health",
                new dev.erst.gridgrind.contract.selector.NamedRangeSelector.All(),
                new InspectionAnalysisQuery.AnalyzeNamedRangeHealth()));
    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            writtenSaveAs(workbookPath.toString(), workbookPath),
            List.of(),
            List.of(),
            List.of(
                new WorkbookInspectionResult.WorkbookSummaryResult(
                    "summary",
                    new WorkbookSummary.WithSheets(
                        1, List.of("Budget"), "Budget", List.of("Budget"), 1, false)),
                new SheetInspectionResult.CellsResult(
                    "cells", "Budget", List.of(textCell("A1", "Report"))),
                new SheetInspectionResult.DataValidationsResult(
                    "data-validations",
                    "Budget",
                    List.of(
                        new DataValidationEntryReport.Unsupported(
                            List.of("A2:A5"), "MISSING_FORMULA", "Formula is missing"))),
                new SheetInspectionResult.HyperlinksResult(
                    "hyperlinks",
                    "Budget",
                    List.of(
                        new CellHyperlinkReport(
                            "A1", new HyperlinkTarget.Url("https://example.com/report")))),
                new WorkbookAnalysisResult.DataValidationHealthResult(
                    "data-validation-health",
                    new DataValidationHealthReport(
                        1,
                        new AnalysisSummaryReport(1, 1, 0, 0),
                        List.of(
                            new AnalysisFindingReport(
                                AnalysisFindingCode.DATA_VALIDATION_UNSUPPORTED_RULE,
                                AnalysisSeverity.WARNING,
                                "Unsupported data-validation rule",
                                "Formula is missing",
                                new AnalysisLocationReport.Range("Budget", "A2:A5"),
                                List.of("A2:A5"))))),
                new WorkbookAnalysisResult.NamedRangeHealthResult(
                    "named-range-health",
                    new NamedRangeHealthReport(
                        1, new AnalysisSummaryReport(0, 0, 0, 0), List.of()))));

    assertDoesNotThrow(
        () -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }

  @Test
  void acceptsNormalizedFileHyperlinkTargets(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");

    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            writtenSaveAs(workbookPath.toString(), workbookPath),
            List.of(),
            List.of(),
            List.of(
                new SheetInspectionResult.HyperlinksResult(
                    "hyperlinks",
                    "Budget",
                    List.of(
                        new CellHyperlinkReport(
                            "A1", new HyperlinkTarget.File("/tmp/report.xlsx"))))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsWorkbookShapeWithNamedRanges() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("B4", ExcelCellValue.number(61.0));
      workbook
          .names()
          .setNamedRange(
              new dev.erst.gridgrind.excel.ExcelNamedRangeDefinition(
                  "BudgetTotal",
                  new dev.erst.gridgrind.excel.ExcelNamedRangeScope.WorkbookScope(),
                  dev.erst.gridgrind.excel.ExcelNamedRangeTarget.range("Budget", "B4")));

      assertDoesNotThrow(() -> WorkbookInvariantChecks.requireWorkbookShape(workbook));
    }
  }

  @Test
  void acceptsWorkbookShapeWithB1SheetState() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Alpha");
      workbook.getOrCreateSheet("Beta");
      workbook.sheets().setActiveSheet("Beta");
      workbook.sheets().setSelectedSheets(List.of("Alpha", "Beta"));
      workbook.sheets().setSheetVisibility("Alpha", ExcelSheetVisibility.HIDDEN);
      workbook.sheets().setSheetProtection("Beta", protectionSettings());

      assertDoesNotThrow(() -> WorkbookInvariantChecks.requireWorkbookShape(workbook));
    }
  }

  @Test
  void acceptsResponseShapeWithProtectedSheetSummary(@TempDir Path tempDirectory)
      throws IOException {
    Path workbookPath = tempDirectory.resolve("result.xlsx");
    Files.writeString(workbookPath, "seed");

    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            writtenSaveAs("result.xlsx", workbookPath),
            List.of(),
            List.of(),
            List.of(
                new SheetInspectionResult.SheetSummaryResult(
                    "sheet",
                    new SheetSummaryReport(
                        "Budget",
                        ExcelSheetVisibility.VERY_HIDDEN,
                        new SheetProtectionReport.Protected(protocolProtectionSettings()),
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

    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            writtenSaveAs("result.xlsx", workbookPath),
            List.of(),
            List.of(),
            List.of(
                new SheetInspectionResult.AutofiltersResult(
                    "autofilters",
                    "Budget",
                    List.of(
                        new AutofilterEntryReport.SheetOwned("E1:F3"),
                        new AutofilterEntryReport.TableOwned("A1:C4", "BudgetTable"))),
                new WorkbookAssetInspectionResult.TablesResult(
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
                new WorkbookAnalysisResult.AutofilterHealthResult(
                    "autofilter-health",
                    new AutofilterHealthReport(
                        2, new AnalysisSummaryReport(0, 0, 0, 0), List.of())),
                new WorkbookAnalysisResult.TableHealthResult(
                    "table-health",
                    new TableHealthReport(1, new AnalysisSummaryReport(0, 0, 0, 0), List.of()))));

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
                new SheetIntrospectionQuery.GetAutofilters()),
            inspect(
                "tables", new TableSelector.All(), new WorkbookAssetIntrospectionQuery.GetTables()),
            inspect(
                "autofilter-health",
                new SheetSelector.All(),
                new InspectionAnalysisQuery.AnalyzeAutofilterHealth()),
            inspect(
                "table-health",
                new TableSelector.All(),
                new InspectionAnalysisQuery.AnalyzeTableHealth()));
    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            writtenSaveAs(workbookPath.toString(), workbookPath),
            List.of(),
            List.of(),
            List.of(
                new SheetInspectionResult.AutofiltersResult(
                    "autofilters",
                    "Budget",
                    List.of(new AutofilterEntryReport.SheetOwned("E1:F3"))),
                new WorkbookAssetInspectionResult.TablesResult(
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
                new WorkbookAnalysisResult.AutofilterHealthResult(
                    "autofilter-health",
                    new AutofilterHealthReport(
                        1, new AnalysisSummaryReport(0, 0, 0, 0), List.of())),
                new WorkbookAnalysisResult.TableHealthResult(
                    "table-health",
                    new TableHealthReport(1, new AnalysisSummaryReport(0, 0, 0, 0), List.of()))));

    assertDoesNotThrow(
        () -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }

  @Test
  void acceptsSuccessResponsesWithPivotReads(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("pivot.xlsx");
    Files.writeString(workbookPath, "seed");

    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            writtenSaveAs("pivot.xlsx", workbookPath),
            List.of(),
            List.of(),
            List.of(
                new WorkbookAssetInspectionResult.PivotTablesResult(
                    "pivots", List.of(pivotReport())),
                new WorkbookAnalysisResult.PivotTableHealthResult(
                    "pivot-health",
                    new PivotTableHealthReport(
                        1, new AnalysisSummaryReport(0, 0, 0, 0), List.of()))));

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
                new WorkbookAssetIntrospectionQuery.GetPivotTables()),
            inspect(
                "pivot-health",
                new PivotTableSelector.All(),
                new InspectionAnalysisQuery.AnalyzePivotTableHealth()));
    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            writtenSaveAs(workbookPath.toString(), workbookPath),
            List.of(),
            List.of(),
            List.of(
                new WorkbookAssetInspectionResult.PivotTablesResult(
                    "pivots", List.of(pivotReport())),
                new WorkbookAnalysisResult.PivotTableHealthResult(
                    "pivot-health",
                    new PivotTableHealthReport(
                        1, new AnalysisSummaryReport(0, 0, 0, 0), List.of()))));

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
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            List.of(),
            List.of(
                assertThat(
                    "assert-total",
                    new SheetSelector.ByName("Budget"),
                    new AnalysisAssertion.AnalysisMaxSeverity(
                        new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                        AnalysisSeverity.ERROR))),
            List.of(
                inspect(
                    "sheet",
                    new SheetSelector.ByName("Budget"),
                    new SheetIntrospectionQuery.GetSheetSummary())));
    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            writtenSaveAs(workbookPath.toString(), workbookPath),
            List.of(),
            List.of(new AssertionResult.Passed("assert-total", "EXPECT_ANALYSIS_MAX_SEVERITY")),
            List.of(
                new SheetInspectionResult.SheetSummaryResult(
                    "sheet",
                    new SheetSummaryReport(
                        "Budget",
                        ExcelSheetVisibility.VISIBLE,
                        new SheetProtectionReport.Unprotected(),
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

    CommentReport anchoredComment =
        new CommentReport(
            "Review",
            "GridGrind",
            true,
            Optional.of(
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
                            false)))),
            Optional.of(new CommentAnchorReport(1, 2, 4, 6)));
    CellStyleReport advancedStyle =
        new CellStyleReport(
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
            CellFillReport.gradient(
                CellGradientFillReport.linear(
                    45.0d,
                    List.of(
                        new CellGradientStopReport(0.0d, rgb("#1F497D")),
                        new CellGradientStopReport(1.0d, themed(4, 0.45d))))),
            new CellBorderReport(
                new CellBorderSideReport.None(),
                new CellBorderSideReport.None(),
                new CellBorderSideReport.None(),
                new CellBorderSideReport.None()),
            new CellProtectionReport(true, false));

    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            writtenSaveAs("advanced.xlsx", workbookPath),
            List.of(),
            List.of(),
            List.of(
                new WorkbookInspectionResult.WorkbookProtectionResult(
                    "workbook-protection",
                    new WorkbookProtectionReport(true, false, true, true, false)),
                new SheetInspectionResult.CellsResult(
                    "cells",
                    "Budget",
                    List.of(
                        new dev.erst.gridgrind.contract.dto.CellReport.TextReport(
                            "A1",
                            java.util.Optional.empty(),
                            java.util.Optional.of(advancedStyle),
                            java.util.Optional.empty(),
                            java.util.Optional.of(anchoredComment),
                            java.util.Optional.of("Review"),
                            java.util.Optional.of(
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
                                            false))))))),
                new SheetInspectionResult.CommentsResult(
                    "comments", "Budget", List.of(new CellCommentReport("A1", anchoredComment))),
                new SheetInspectionResult.PrintLayoutResult(
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
                new SheetInspectionResult.AutofiltersResult(
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
                            Optional.of(
                                new AutofilterSortStateReport(
                                    "A2:C5",
                                    false,
                                    false,
                                    java.util.Optional.empty(),
                                    List.of(
                                        new AutofilterSortConditionReport.Value(
                                            "B2:B5", true))))))),
                new WorkbookAssetInspectionResult.TablesResult(
                    "tables",
                    List.of(
                        new TableEntryReport(
                            "BudgetTable",
                            "Budget",
                            "A1:C5",
                            new TableEntryReport.Structure(
                                1,
                                1,
                                List.of("Item", "Status", "Owner"),
                                List.of(
                                    new TableColumnReport(
                                        1L,
                                        "Item",
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.empty()),
                                    new TableColumnReport(
                                        2L,
                                        "Status",
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.of("sum"),
                                        Optional.empty()),
                                    new TableColumnReport(
                                        3L,
                                        "Owner",
                                        Optional.of("owner_unique"),
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.of("CONCAT([@Owner],\"-\",[@Status])")))),
                            new TableStyleReport.Named(
                                "TableStyleMedium2", false, false, true, false),
                            new TableEntryReport.Behavior(true, true, false, true),
                            new TableEntryReport.Presentation(
                                Optional.of("Team queue"),
                                Optional.of("HeaderStyle"),
                                Optional.of("DataStyle"),
                                Optional.of("TotalsStyle")))))));

    assertDoesNotThrow(() -> WorkbookInvariantChecks.requireResponseShape(response));
  }

  @Test
  void acceptsSuccessResponsesWithDrawingReadShapes(@TempDir Path tempDirectory)
      throws IOException {
    Path workbookPath = tempDirectory.resolve("drawing.xlsx");
    Files.writeString(workbookPath, "seed");

    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            writtenSaveAs("drawing.xlsx", workbookPath),
            List.of(),
            List.of(),
            List.of(
                new WorkbookAssetInspectionResult.DrawingObjectsResult(
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
                new WorkbookAssetInspectionResult.DrawingObjectPayloadResult(
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
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
            inspect(
                "drawing-payload",
                new DrawingObjectSelector.ByName("Ops", "OpsPicture"),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload()));
    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            writtenSaveAs(workbookPath.toString(), workbookPath),
            List.of(),
            List.of(),
            List.of(
                new WorkbookAssetInspectionResult.DrawingObjectsResult(
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
                new WorkbookAssetInspectionResult.DrawingObjectPayloadResult(
                    "drawing-payload",
                    "Ops",
                    new DrawingObjectPayloadReport.Picture(
                        "OpsPicture",
                        ExcelPictureFormat.PNG,
                        "image/png",
                        "OpsPicture.png",
                        "abc123",
                        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC",
                        null))));

    assertDoesNotThrow(
        () -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }

  @Test
  void acceptsSuccessResponsesWithChartReadShapes(@TempDir Path tempDirectory) throws IOException {
    Path workbookPath = tempDirectory.resolve("chart.xlsx");
    Files.writeString(workbookPath, "seed");

    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            writtenSaveAs("chart.xlsx", workbookPath),
            List.of(),
            List.of(),
            List.of(
                new WorkbookAssetInspectionResult.DrawingObjectsResult(
                    "drawing-objects",
                    "Ops",
                    List.of(
                        new DrawingObjectReport.Chart(
                            "OpsChart", twoCellAnchor(), true, List.of("BAR"), "Roadmap"))),
                new WorkbookAssetInspectionResult.ChartsResult(
                    "charts", "Ops", List.of(chartReport()))));

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
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
            inspect(
                "charts",
                new ChartSelector.AllOnSheet("Ops"),
                new WorkbookAssetIntrospectionQuery.GetCharts()));
    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            writtenSaveAs(workbookPath.toString(), workbookPath),
            List.of(),
            List.of(),
            List.of(
                new WorkbookAssetInspectionResult.DrawingObjectsResult(
                    "drawing-objects",
                    "Ops",
                    List.of(
                        new DrawingObjectReport.Chart(
                            "OpsChart", twoCellAnchor(), true, List.of("BAR"), "Roadmap"))),
                new WorkbookAssetInspectionResult.ChartsResult(
                    "charts", "Ops", List.of(chartReport()))));

    assertDoesNotThrow(
        () -> WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response));
  }

  @Test
  void acceptsWorkbookShapeWithCharts() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Ops").cells().setCell("A1", ExcelCellValue.text("Month"));
      workbook.getOrCreateSheet("Ops").cells().setCell("B1", ExcelCellValue.text("Actual"));
      workbook.getOrCreateSheet("Ops").cells().setCell("A2", ExcelCellValue.text("Jan"));
      workbook.getOrCreateSheet("Ops").cells().setCell("B2", ExcelCellValue.number(12.0d));
      workbook.getOrCreateSheet("Ops").cells().setCell("A3", ExcelCellValue.text("Feb"));
      workbook.getOrCreateSheet("Ops").cells().setCell("B3", ExcelCellValue.number(18.0d));
      workbook.getOrCreateSheet("Ops").cells().setCell("A4", ExcelCellValue.text("Mar"));
      workbook.getOrCreateSheet("Ops").cells().setCell("B4", ExcelCellValue.number(15.0d));
      new WorkbookExecutionEngine()
          .apply(
              workbook,
              List.of(
                  new WorkbookDrawingCommand.SetChart(
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
                                  Optional.empty(),
                                  Optional.empty(),
                                  excelCategoryAxes(),
                                  List.of(
                                      new ExcelChartDefinition.Series(
                                          new ExcelChartDefinition.Title.Text("Actual"),
                                          new ExcelChartDefinition.DataSource.Reference(
                                              "Ops!$A$2:$A$4"),
                                          new ExcelChartDefinition.DataSource.Reference(
                                              "Ops!$B$2:$B$4"),
                                          Optional.empty(),
                                          Optional.empty(),
                                          Optional.empty(),
                                          Optional.empty()))))))));

      assertDoesNotThrow(() -> WorkbookInvariantChecks.requireWorkbookShape(workbook));
    }
  }

  @Test
  void acceptsWorkbookShapeWithPivots() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget").cells().setCell("A1", ExcelCellValue.text("Month"));
      workbook.getOrCreateSheet("Budget").cells().setCell("B1", ExcelCellValue.text("Actual"));
      workbook.getOrCreateSheet("Budget").cells().setCell("A2", ExcelCellValue.text("Jan"));
      workbook.getOrCreateSheet("Budget").cells().setCell("B2", ExcelCellValue.number(12.0d));
      workbook.getOrCreateSheet("Budget").cells().setCell("A3", ExcelCellValue.text("Feb"));
      workbook.getOrCreateSheet("Budget").cells().setCell("B3", ExcelCellValue.number(18.0d));
      workbook.getOrCreateSheet("Budget").cells().setCell("A4", ExcelCellValue.text("Mar"));
      workbook.getOrCreateSheet("Budget").cells().setCell("B4", ExcelCellValue.number(15.0d));
      workbook.getOrCreateSheet("Pivot");
      new WorkbookExecutionEngine()
          .apply(
              workbook,
              List.of(
                  new WorkbookTabularCommand.SetPivotTable(
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
                                  Optional.empty()))))));

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
                new dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput(
                    java.util.Optional.of("GridGrind-2026"))),
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new WorkbookIntrospectionQuery.GetPackageSecurity()));
    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            new WorkbookResultPersistence.PersistenceOutcome.NotSaved(),
            List.of(),
            List.of(),
            List.of(
                new WorkbookInspectionResult.PackageSecurityResult(
                    "security",
                    new dev.erst.gridgrind.contract.dto.OoxmlPackageSecurityReport(
                        new dev.erst.gridgrind.contract.dto.OoxmlEncryptionReport.Encrypted(
                            ExcelOoxmlEncryptionMode.AGILE,
                            ExcelOoxmlCipherAlgorithm.AES_256,
                            ExcelOoxmlHashAlgorithm.SHA_512,
                            ExcelOoxmlChainingMode.CBC,
                            256,
                            16,
                            100000),
                        List.of(
                            new dev.erst.gridgrind.contract.dto.OoxmlSignatureReport(
                                "/_xmlsignatures/sig1.xml",
                                Optional.of("CN=GridGrind Signing"),
                                Optional.of("CN=GridGrind Signing"),
                                Optional.of("01"),
                                ExcelOoxmlSignatureState.VALID))))));

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
                new WorkbookIntrospectionQuery.GetCustomXmlMappings()),
            inspect(
                "custom-xml-export",
                new WorkbookSelector.Current(),
                new WorkbookIntrospectionQuery.ExportCustomXmlMapping(
                    new CustomXmlMappingLocator(1L, "BudgetMap"), true, "UTF-8")),
            inspect(
                "array-formulas",
                new SheetSelector.ByName("Ops"),
                new SheetIntrospectionQuery.GetArrayFormulas()),
            inspect(
                "drawing-objects",
                new DrawingObjectSelector.AllOnSheet("Ops"),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects()));
    CustomXmlMappingReport mapping = customXmlMappingReport();
    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V2,
            new WorkbookResultPersistence.PersistenceOutcome.NotSaved(),
            List.of(),
            List.of(),
            List.of(
                new WorkbookInspectionResult.CustomXmlMappingsResult(
                    "custom-xml-mappings", List.of(mapping)),
                new WorkbookInspectionResult.CustomXmlExportResult(
                    "custom-xml-export",
                    new CustomXmlExportReport(
                        mapping,
                        "UTF-8",
                        true,
                        "<BudgetMap><Owner>Ada Lovelace</Owner></BudgetMap>")),
                new SheetInspectionResult.ArrayFormulasResult(
                    "array-formulas",
                    List.of(new ArrayFormulaReport("Ops", "D2:D4", "D2", "B2:B4*C2:C4", false))),
                new WorkbookAssetInspectionResult.DrawingObjectsResult(
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
                Optional.empty(),
                Optional.empty(),
                chartCategoryAxes(),
                List.of(
                    new ChartReport.Series(
                        new ChartReport.Title.Text("Actual"),
                        new ChartReport.DataSource.StringReference(
                            "Ops!$A$2:$A$4", List.of("Jan", "Feb", "Mar")),
                        new ChartReport.DataSource.NumericReference(
                            "Ops!$B$2:$B$4", Optional.of("General"), List.of("12", "18", "15")),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty())))));
  }

  private static CustomXmlMappingReport customXmlMappingReport() {
    return new CustomXmlMappingReport(
        1L,
        "BudgetMap",
        "BudgetMap",
        "schema-1",
        new CustomXmlMappingReport.Settings(true, true, false, true, true),
        new CustomXmlMappingReport.Schema(
            "urn:gridgrind:budget", "en-US", "budget-map.xsd", "<xs:schema/>"),
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
        Optional.of(
            new DrawingObjectReport.SignatureSetup(
                Optional.of("sig-setup-01"),
                Optional.of(false),
                Optional.of("Review the budget before signing."),
                Optional.of("Ada Lovelace"),
                Optional.of("Finance"),
                Optional.of("ada@example.com"))),
        Optional.of(
            new DrawingObjectReport.SignaturePreview(
                ExcelPictureFormat.PNG,
                "image/png",
                128L,
                Optional.of("0123456789abcdef"),
                Optional.of(320),
                Optional.of(120))));
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

  private static WorkbookResultPersistence.PersistenceOutcome.SavedAs writtenSaveAs(
      String requestedPath, Path workbookPath) {
    return new WorkbookResultPersistence.PersistenceOutcome.SavedAs(
        requestedPath, new WorkbookResultPersistence.WriteResult.Written(workbookPath.toString()));
  }

  private static WorkbookResultPersistence.PersistenceOutcome.SavedAs notWrittenSaveAs(
      String requestedPath) {
    return new WorkbookResultPersistence.PersistenceOutcome.SavedAs(
        requestedPath, new WorkbookResultPersistence.WriteResult.NotWritten());
  }

  @SafeVarargs
  private static WorkbookPlan saveAsRequest(Path workbookPath, InspectionStep... inspections) {
    return ProtocolStepSupport.request(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.SaveAs(
            workbookPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
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
                2,
                "Actual",
                ExcelPivotDataConsolidateFunction.SUM,
                "Total Actual",
                Optional.of("General"))),
        false);
  }

  private static CellStyleReport defaultStyle() {
    return new CellStyleReport(
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
        CellFillReport.pattern(ExcelFillPattern.NONE),
        new CellBorderReport(
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None(),
            new CellBorderSideReport.None()),
        new CellProtectionReport(true, false));
  }

  private static dev.erst.gridgrind.contract.dto.CellReport.TextReport textCell(
      String address, String value) {
    return new dev.erst.gridgrind.contract.dto.CellReport.TextReport(
        address,
        java.util.Optional.of(value),
        java.util.Optional.of(defaultStyle()),
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.of(value),
        java.util.Optional.empty());
  }

  private static CellColorReport rgb(String rgb) {
    return CellColorReport.rgb(rgb);
  }

  private static CellColorReport indexed(int indexed) {
    return CellColorReport.indexed(indexed);
  }

  private static CellColorReport themed(int theme, double tint) {
    return CellColorReport.theme(theme, tint);
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
