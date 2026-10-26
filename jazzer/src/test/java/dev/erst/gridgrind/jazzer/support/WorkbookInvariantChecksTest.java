package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.assertThat;
import static dev.erst.gridgrind.jazzer.support.ProtocolStepSupport.inspect;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlDataBindingReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlExportReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedCellReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlLinkedTableReport;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingLocator;
import dev.erst.gridgrind.core.model.customxml.CustomXmlMappingReport;
import dev.erst.gridgrind.core.model.drawing.CellCommentReport;
import dev.erst.gridgrind.core.model.drawing.CommentAnchorReport;
import dev.erst.gridgrind.core.model.drawing.CommentReport;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectPayloadReport;
import dev.erst.gridgrind.core.model.drawing.DrawingObjectReport;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaReport;
import dev.erst.gridgrind.core.model.formula.FormulaHealthReport;
import dev.erst.gridgrind.core.model.formula.FormulaPatternReport;
import dev.erst.gridgrind.core.model.formula.FormulaSurfaceReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeHealthReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeSurfaceEntryReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeSurfaceReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.formula.SheetFormulaSurfaceReport;
import dev.erst.gridgrind.core.model.layout.ColumnLayoutReport;
import dev.erst.gridgrind.core.model.layout.MergedRegionReport;
import dev.erst.gridgrind.core.model.layout.PaneReport;
import dev.erst.gridgrind.core.model.layout.RowLayoutReport;
import dev.erst.gridgrind.core.model.layout.SheetLayoutReport;
import dev.erst.gridgrind.core.model.layout.WindowDimensionsReport;
import dev.erst.gridgrind.core.model.layout.WindowReport;
import dev.erst.gridgrind.core.model.layout.WindowRowReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableHealthReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintAreaReport;
import dev.erst.gridgrind.core.model.print.PrintLayoutReport;
import dev.erst.gridgrind.core.model.print.PrintMarginsReport;
import dev.erst.gridgrind.core.model.print.PrintScalingReport;
import dev.erst.gridgrind.core.model.print.PrintSetupReport;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsReport;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsReport;
import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellGradientFillReport;
import dev.erst.gridgrind.core.model.style.CellGradientStopReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.RichTextRunReport;
import dev.erst.gridgrind.core.model.table.AutofilterEntryReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnReport;
import dev.erst.gridgrind.core.model.table.AutofilterFilterCriterionReport;
import dev.erst.gridgrind.core.model.table.AutofilterHealthReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortConditionReport;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateReport;
import dev.erst.gridgrind.core.model.table.SchemaColumnReport;
import dev.erst.gridgrind.core.model.table.SheetSchemaReport;
import dev.erst.gridgrind.core.model.table.TableColumnReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableHealthReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationHealthReport;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleInput;
import dev.erst.gridgrind.core.model.value.CellHyperlinkReport;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.GridGrindProtocolVersion;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.SheetInspectionResult;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookAnalysisResult;
import dev.erst.gridgrind.core.plan.WorkbookAssetInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.plan.WorkbookSurfaceInspectionResult;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.GridGrindWarningCode;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces;
import dev.erst.gridgrind.core.problem.RequestWarning;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.result.AnalysisSummaryReport;
import dev.erst.gridgrind.core.result.SheetSummaryReport;
import dev.erst.gridgrind.core.result.TypeCountReport;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.result.WorkbookResultPersistence;
import dev.erst.gridgrind.core.result.WorkbookResults;
import dev.erst.gridgrind.core.result.WorkbookSummary;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisCrosses;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisKind;
import dev.erst.gridgrind.core.vocabulary.ExcelChartAxisPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarDirection;
import dev.erst.gridgrind.core.vocabulary.ExcelChartBarGrouping;
import dev.erst.gridgrind.core.vocabulary.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.core.vocabulary.ExcelChartLegendPosition;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingShapeKind;
import dev.erst.gridgrind.core.vocabulary.ExcelEmbeddedObjectPackagingKind;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlChainingMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlCipherAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlEncryptionMode;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlHashAlgorithm;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlSignatureState;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.NamedRangeBackingKind;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.excel.*;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelChartDefinition;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingAnchor;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingMarker;
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
            GridGrindProtocolVersion.V3,
            publishedSaveAs("result.xlsx", workbookPath),
            List.of(
                new RequestWarning(
                    GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
                    1,
                    new StepId("step-01-set-cell"),
                    "SET_CELL",
                    "Formula references same-request sheet names with spaces.")),
            List.of(
                new AssertionResult.Passed(
                    new StepId("assert-total"), "EXPECT_NAMED_RANGE_PRESENT")),
            List.of(
                new WorkbookInspectionResult.WorkbookSummaryResult(
                    new StepId("summary"),
                    new WorkbookSummary.WithSheets(
                        1,
                        List.of(new SheetLabel("Budget")),
                        new SheetLabel("Budget"),
                        List.of(new SheetLabel("Budget")),
                        1,
                        false)),
                new WorkbookInspectionResult.NamedRangesResult(
                    new StepId("ranges"),
                    List.of(
                        new NamedRangeReport.RangeReport(
                            "BudgetTotal",
                            new NamedRangeScope.Workbook(),
                            "Budget!$B$4",
                            NamedRangeTarget.range(new SheetName("Budget"), new RangeRef("B4"))))),
                new SheetInspectionResult.CellsResult(
                    new StepId("cells"),
                    new SheetLabel("Budget"),
                    List.of(
                        new dev.erst.gridgrind.core.model.value.CellReport.TextReport(
                            new CellAddress("A1"),
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
                    new StepId("window"),
                    new WindowReport.Dense(
                        new SheetLabel("Budget"),
                        new CellAddress("A1"),
                        new WindowDimensionsReport(1, 1),
                        List.of(
                            new WindowRowReport(
                                0,
                                List.of(
                                    new dev.erst.gridgrind.core.model.value.CellReport.TextReport(
                                        new CellAddress("A1"),
                                        java.util.Optional.empty(),
                                        java.util.Optional.of(style),
                                        java.util.Optional.empty(),
                                        java.util.Optional.empty(),
                                        java.util.Optional.of("Report"),
                                        java.util.Optional.empty())))))),
                new SheetInspectionResult.MergedRegionsResult(
                    new StepId("merged"),
                    new SheetLabel("Budget"),
                    List.of(new MergedRegionReport(new RangeLabel("A1:B1")))),
                new SheetInspectionResult.HyperlinksResult(
                    new StepId("hyperlinks"),
                    new SheetLabel("Budget"),
                    List.of(
                        new CellHyperlinkReport(
                            new CellAddress("A1"),
                            new HyperlinkTarget.Url("https://example.com/report")))),
                new SheetInspectionResult.CommentsResult(
                    new StepId("comments"),
                    new SheetLabel("Budget"),
                    List.of(
                        new CellCommentReport(
                            new CellAddress("A1"),
                            new CommentReport("Review", "GridGrind", true)))),
                new SheetInspectionResult.SheetLayoutResult(
                    new StepId("layout"),
                    new SheetLayoutReport(
                        new SheetLabel("Budget"),
                        new PaneReport.Frozen(1, 1, 1, 1),
                        125,
                        dev.erst.gridgrind.core.model.layout.SheetPresentationReport.defaults(),
                        List.of(new ColumnLayoutReport(0, 12.5, false, 0, false)),
                        List.of(new RowLayoutReport(0, 18.0, false, 0, false)))),
                new SheetInspectionResult.PrintLayoutResult(
                    new StepId("print-layout"),
                    new PrintLayoutReport(
                        new SheetLabel("Budget"),
                        new PrintAreaReport.Range(new RangeLabel("A1:B20")),
                        ExcelPrintOrientation.LANDSCAPE,
                        new PrintScalingReport.Fit(1, 0),
                        new PrintTitleRowsReport.Band(0, 0),
                        new PrintTitleColumnsReport.Band(0, 0),
                        new HeaderFooterTextReport("Budget", "", ""),
                        new HeaderFooterTextReport("", "Page &P", ""))),
                new SheetInspectionResult.DataValidationsResult(
                    new StepId("data-validations"),
                    new SheetLabel("Budget"),
                    List.of(
                        new DataValidationEntryReport.Supported(
                            List.of(new RangeLabel("A2:A5")),
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
                    new StepId("formula-surface"),
                    new FormulaSurfaceReport(
                        1,
                        List.of(
                            new SheetFormulaSurfaceReport(
                                new SheetLabel("Budget"),
                                1,
                                1,
                                List.of(
                                    new FormulaPatternReport(
                                        "SUM(B2:B3)", 1, List.of(new CellAddress("B4")))))))),
                new WorkbookSurfaceInspectionResult.SheetSchemaResult(
                    new StepId("schema"),
                    new SheetSchemaReport(
                        new SheetLabel("Budget"),
                        new CellAddress("A1"),
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
                    new StepId("named-range-surface"),
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
                    new StepId("formula-health"),
                    new FormulaHealthReport(
                        1,
                        new AnalysisSummaryReport(1, 0, 0, 1),
                        List.of(
                            new AnalysisFindingReport(
                                AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                                AnalysisSeverity.INFO,
                                "Volatile formula",
                                "Formula uses NOW().",
                                new AnalysisLocationReport.Cell(new SheetLabel("Budget"), "B4"),
                                List.of("NOW()"))))),
                new WorkbookAnalysisResult.DataValidationHealthResult(
                    new StepId("data-validation-health"),
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
            GridGrindProtocolVersion.V3,
            notAttemptedSaveAs(workbookPath.toString()),
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
                new CellSelector.ByAddresses(
                    new SheetName("Budget"), List.of(new CellAddress("A1"))),
                new SheetIntrospectionQuery.GetCells()),
            inspect(
                "data-validations",
                new RangeSelector.AllOnSheet(new SheetName("Budget")),
                new SheetIntrospectionQuery.GetDataValidations()),
            inspect(
                "hyperlinks",
                new CellSelector.AllUsedInSheet(new SheetName("Budget")),
                new SheetIntrospectionQuery.GetHyperlinks()),
            inspect(
                "data-validation-health",
                new SheetSelector.All(),
                new InspectionAnalysisQuery.AnalyzeDataValidationHealth()),
            inspect(
                "named-range-health",
                new dev.erst.gridgrind.core.plan.NamedRangeSelector.All(),
                new InspectionAnalysisQuery.AnalyzeNamedRangeHealth()));
    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V3,
            publishedSaveAs(workbookPath.toString(), workbookPath),
            List.of(),
            List.of(),
            List.of(
                new WorkbookInspectionResult.WorkbookSummaryResult(
                    new StepId("summary"),
                    new WorkbookSummary.WithSheets(
                        1,
                        List.of(new SheetLabel("Budget")),
                        new SheetLabel("Budget"),
                        List.of(new SheetLabel("Budget")),
                        1,
                        false)),
                new SheetInspectionResult.CellsResult(
                    new StepId("cells"),
                    new SheetLabel("Budget"),
                    List.of(textCell("A1", "Report"))),
                new SheetInspectionResult.DataValidationsResult(
                    new StepId("data-validations"),
                    new SheetLabel("Budget"),
                    List.of(
                        new DataValidationEntryReport.Unsupported(
                            List.of(new RangeLabel("A2:A5")),
                            "MISSING_FORMULA",
                            "Formula is missing"))),
                new SheetInspectionResult.HyperlinksResult(
                    new StepId("hyperlinks"),
                    new SheetLabel("Budget"),
                    List.of(
                        new CellHyperlinkReport(
                            new CellAddress("A1"),
                            new HyperlinkTarget.Url("https://example.com/report")))),
                new WorkbookAnalysisResult.DataValidationHealthResult(
                    new StepId("data-validation-health"),
                    new DataValidationHealthReport(
                        1,
                        new AnalysisSummaryReport(1, 1, 0, 0),
                        List.of(
                            new AnalysisFindingReport(
                                AnalysisFindingCode.DATA_VALIDATION_UNSUPPORTED_RULE,
                                AnalysisSeverity.WARNING,
                                "Unsupported data-validation rule",
                                "Formula is missing",
                                new AnalysisLocationReport.Range(new SheetLabel("Budget"), "A2:A5"),
                                List.of("A2:A5"))))),
                new WorkbookAnalysisResult.NamedRangeHealthResult(
                    new StepId("named-range-health"),
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
            GridGrindProtocolVersion.V3,
            publishedSaveAs(workbookPath.toString(), workbookPath),
            List.of(),
            List.of(),
            List.of(
                new SheetInspectionResult.HyperlinksResult(
                    new StepId("hyperlinks"),
                    new SheetLabel("Budget"),
                    List.of(
                        new CellHyperlinkReport(
                            new CellAddress("A1"),
                            new HyperlinkTarget.File("/tmp/report.xlsx"))))));

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
            GridGrindProtocolVersion.V3,
            publishedSaveAs("result.xlsx", workbookPath),
            List.of(),
            List.of(),
            List.of(
                new SheetInspectionResult.SheetSummaryResult(
                    new StepId("sheet"),
                    new SheetSummaryReport(
                        new SheetLabel("Budget"),
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
            GridGrindProtocolVersion.V3,
            publishedSaveAs("result.xlsx", workbookPath),
            List.of(),
            List.of(),
            List.of(
                new SheetInspectionResult.AutofiltersResult(
                    new StepId("autofilters"),
                    new SheetLabel("Budget"),
                    List.of(
                        new AutofilterEntryReport.SheetOwned(new RangeLabel("E1:F3")),
                        new AutofilterEntryReport.TableOwned(
                            new RangeLabel("A1:C4"), "BudgetTable"))),
                new WorkbookAssetInspectionResult.TablesResult(
                    new StepId("tables"),
                    List.of(
                        new TableEntryReport(
                            "BudgetTable",
                            new SheetLabel("Budget"),
                            new RangeLabel("A1:C4"),
                            1,
                            0,
                            List.of("Item", "Amount", "Billable"),
                            new TableStyleReport.Named(
                                "TableStyleMedium2", false, false, true, false),
                            true))),
                new WorkbookAnalysisResult.AutofilterHealthResult(
                    new StepId("autofilter-health"),
                    new AutofilterHealthReport(
                        2, new AnalysisSummaryReport(0, 0, 0, 0), List.of())),
                new WorkbookAnalysisResult.TableHealthResult(
                    new StepId("table-health"),
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
                new SheetSelector.ByName(new SheetName("Budget")),
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
            GridGrindProtocolVersion.V3,
            publishedSaveAs(workbookPath.toString(), workbookPath),
            List.of(),
            List.of(),
            List.of(
                new SheetInspectionResult.AutofiltersResult(
                    new StepId("autofilters"),
                    new SheetLabel("Budget"),
                    List.of(new AutofilterEntryReport.SheetOwned(new RangeLabel("E1:F3")))),
                new WorkbookAssetInspectionResult.TablesResult(
                    new StepId("tables"),
                    List.of(
                        new TableEntryReport(
                            "BudgetTable",
                            new SheetLabel("Budget"),
                            new RangeLabel("A1:C4"),
                            1,
                            0,
                            List.of("Item", "Amount", "Billable"),
                            new TableStyleReport.None(),
                            true))),
                new WorkbookAnalysisResult.AutofilterHealthResult(
                    new StepId("autofilter-health"),
                    new AutofilterHealthReport(
                        1, new AnalysisSummaryReport(0, 0, 0, 0), List.of())),
                new WorkbookAnalysisResult.TableHealthResult(
                    new StepId("table-health"),
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
            GridGrindProtocolVersion.V3,
            publishedSaveAs("pivot.xlsx", workbookPath),
            List.of(),
            List.of(),
            List.of(
                new WorkbookAssetInspectionResult.PivotTablesResult(
                    new StepId("pivots"), List.of(pivotReport())),
                new WorkbookAnalysisResult.PivotTableHealthResult(
                    new StepId("pivot-health"),
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
            GridGrindProtocolVersion.V3,
            publishedSaveAs(workbookPath.toString(), workbookPath),
            List.of(),
            List.of(),
            List.of(
                new WorkbookAssetInspectionResult.PivotTablesResult(
                    new StepId("pivots"), List.of(pivotReport())),
                new WorkbookAnalysisResult.PivotTableHealthResult(
                    new StepId("pivot-health"),
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
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new AnalysisAssertion.AnalysisMaxSeverity(
                        new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                        AnalysisSeverity.ERROR))),
            List.of(
                inspect(
                    "sheet",
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new SheetIntrospectionQuery.GetSheetSummary())));
    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V3,
            publishedSaveAs(workbookPath.toString(), workbookPath),
            List.of(),
            List.of(
                new AssertionResult.Passed(
                    new StepId("assert-total"), "EXPECT_ANALYSIS_MAX_SEVERITY")),
            List.of(
                new SheetInspectionResult.SheetSummaryResult(
                    new StepId("sheet"),
                    new SheetSummaryReport(
                        new SheetLabel("Budget"),
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
            GridGrindProtocolVersion.V3,
            publishedSaveAs("advanced.xlsx", workbookPath),
            List.of(),
            List.of(),
            List.of(
                new WorkbookInspectionResult.WorkbookProtectionResult(
                    new StepId("workbook-protection"),
                    new WorkbookProtectionReport(true, false, true, true, false)),
                new SheetInspectionResult.CellsResult(
                    new StepId("cells"),
                    new SheetLabel("Budget"),
                    List.of(
                        new dev.erst.gridgrind.core.model.value.CellReport.TextReport(
                            new CellAddress("A1"),
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
                    new StepId("comments"),
                    new SheetLabel("Budget"),
                    List.of(new CellCommentReport(new CellAddress("A1"), anchoredComment))),
                new SheetInspectionResult.PrintLayoutResult(
                    new StepId("print-layout"),
                    new PrintLayoutReport(
                        new SheetLabel("Budget"),
                        new PrintAreaReport.Range(new RangeLabel("A1:C20")),
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
                    new StepId("autofilters"),
                    new SheetLabel("Budget"),
                    List.of(
                        new AutofilterEntryReport.TableOwned(
                            new RangeLabel("A1:C5"),
                            "BudgetTable",
                            List.of(
                                new AutofilterFilterColumnReport(
                                    1L,
                                    true,
                                    new AutofilterFilterCriterionReport.Values(
                                        List.of("Open", "Closed"), true))),
                            Optional.of(
                                new AutofilterSortStateReport(
                                    new RangeLabel("A2:C5"),
                                    false,
                                    false,
                                    java.util.Optional.empty(),
                                    List.of(
                                        new AutofilterSortConditionReport.Value(
                                            new RangeLabel("B2:B5"), true))))))),
                new WorkbookAssetInspectionResult.TablesResult(
                    new StepId("tables"),
                    List.of(
                        new TableEntryReport(
                            "BudgetTable",
                            new SheetLabel("Budget"),
                            new RangeLabel("A1:C5"),
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
            GridGrindProtocolVersion.V3,
            publishedSaveAs("drawing.xlsx", workbookPath),
            List.of(),
            List.of(),
            List.of(
                new WorkbookAssetInspectionResult.DrawingObjectsResult(
                    new StepId("drawing-objects"),
                    new SheetLabel("Ops"),
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
                    new StepId("drawing-payload"),
                    new SheetLabel("Ops"),
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
                new DrawingObjectSelector.AllOnSheet(new SheetName("Ops")),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
            inspect(
                "drawing-payload",
                new DrawingObjectSelector.ByName(new SheetName("Ops"), "OpsPicture"),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjectPayload()));
    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V3,
            publishedSaveAs(workbookPath.toString(), workbookPath),
            List.of(),
            List.of(),
            List.of(
                new WorkbookAssetInspectionResult.DrawingObjectsResult(
                    new StepId("drawing-objects"),
                    new SheetLabel("Ops"),
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
                    new StepId("drawing-payload"),
                    new SheetLabel("Ops"),
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
            GridGrindProtocolVersion.V3,
            publishedSaveAs("chart.xlsx", workbookPath),
            List.of(),
            List.of(),
            List.of(
                new WorkbookAssetInspectionResult.DrawingObjectsResult(
                    new StepId("drawing-objects"),
                    new SheetLabel("Ops"),
                    List.of(
                        new DrawingObjectReport.Chart(
                            "OpsChart", twoCellAnchor(), true, List.of("BAR"), "Roadmap"))),
                new WorkbookAssetInspectionResult.ChartsResult(
                    new StepId("charts"), new SheetLabel("Ops"), List.of(chartReport()))));

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
                new DrawingObjectSelector.AllOnSheet(new SheetName("Ops")),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
            inspect(
                "charts",
                new ChartSelector.AllOnSheet(new SheetName("Ops")),
                new WorkbookAssetIntrospectionQuery.GetCharts()));
    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V3,
            publishedSaveAs(workbookPath.toString(), workbookPath),
            List.of(),
            List.of(),
            List.of(
                new WorkbookAssetInspectionResult.DrawingObjectsResult(
                    new StepId("drawing-objects"),
                    new SheetLabel("Ops"),
                    List.of(
                        new DrawingObjectReport.Chart(
                            "OpsChart", twoCellAnchor(), true, List.of("BAR"), "Roadmap"))),
                new WorkbookAssetInspectionResult.ChartsResult(
                    new StepId("charts"), new SheetLabel("Ops"), List.of(chartReport()))));

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
                new dev.erst.gridgrind.core.model.security.OoxmlOpenSecurityInput(
                    java.util.Optional.of(
                        new dev.erst.gridgrind.core.source.SecretReference(
                            "source-open-password")))),
            inspect(
                "security",
                new WorkbookSelector.Current(),
                new WorkbookIntrospectionQuery.GetPackageSecurity()));
    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V3,
            new WorkbookResultPersistence.PersistenceOutcome.NotSaved(),
            List.of(),
            List.of(),
            List.of(
                new WorkbookInspectionResult.PackageSecurityResult(
                    new StepId("security"),
                    new dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport(
                        new dev.erst.gridgrind.core.model.security.OoxmlEncryptionReport.Encrypted(
                            ExcelOoxmlEncryptionMode.AGILE,
                            ExcelOoxmlCipherAlgorithm.AES_256,
                            ExcelOoxmlHashAlgorithm.SHA_512,
                            ExcelOoxmlChainingMode.CBC,
                            256,
                            16,
                            100000),
                        List.of(
                            new dev.erst.gridgrind.core.model.security.OoxmlSignatureReport(
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
                new SheetSelector.ByName(new SheetName("Ops")),
                new SheetIntrospectionQuery.GetArrayFormulas()),
            inspect(
                "drawing-objects",
                new DrawingObjectSelector.AllOnSheet(new SheetName("Ops")),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects()));
    CustomXmlMappingReport mapping = customXmlMappingReport();
    WorkbookResult.Success response =
        WorkbookResults.success(
            GridGrindProtocolVersion.V3,
            new WorkbookResultPersistence.PersistenceOutcome.NotSaved(),
            List.of(),
            List.of(),
            List.of(
                new WorkbookInspectionResult.CustomXmlMappingsResult(
                    new StepId("custom-xml-mappings"), List.of(mapping)),
                new WorkbookInspectionResult.CustomXmlExportResult(
                    new StepId("custom-xml-export"),
                    new CustomXmlExportReport(
                        mapping,
                        "UTF-8",
                        true,
                        "<BudgetMap><Owner>Ada Lovelace</Owner></BudgetMap>")),
                new SheetInspectionResult.ArrayFormulasResult(
                    new StepId("array-formulas"),
                    List.of(
                        new ArrayFormulaReport(
                            new SheetLabel("Ops"),
                            new RangeLabel("D2:D4"),
                            new CellAddress("D2"),
                            "B2:B4*C2:C4",
                            false))),
                new WorkbookAssetInspectionResult.DrawingObjectsResult(
                    new StepId("drawing-objects"),
                    new SheetLabel("Ops"),
                    List.of(signatureLineDrawingObjectReport()))));

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
        List.of(
            new CustomXmlLinkedCellReport(
                new SheetLabel("Ops"), new CellAddress("A2"), "/BudgetMap/Owner[1]", "string")),
        List.of(
            new CustomXmlLinkedTableReport(
                new SheetLabel("Ops"),
                "BudgetTable",
                "BudgetTable",
                new RangeLabel("A1:B4"),
                "/BudgetMap/Rows")));
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

  private static WorkbookResultPersistence.PersistenceOutcome.SavedAs publishedSaveAs(
      String requestedPath, Path workbookPath) {
    return new WorkbookResultPersistence.PersistenceOutcome.SavedAs(
        requestedPath,
        new WorkbookResultPersistence.PublicationOutcome.Published(
            workbookPath.toString(),
            "0".repeat(64),
            0,
            new WorkbookResultPersistence.PublicationOutcome.StagedArtifactVerification(),
            new WorkbookResultPersistence.PublicationOutcome.DurabilityEvidence
                .FileSyncedDirectoryUnestablished()));
  }

  private static WorkbookResultPersistence.PersistenceOutcome.SavedAs notAttemptedSaveAs(
      String requestedPath) {
    return new WorkbookResultPersistence.PersistenceOutcome.SavedAs(
        requestedPath, new WorkbookResultPersistence.PublicationOutcome.NotAttempted());
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
        new SheetName("Budget"),
        new PivotTableReport.Anchor(new CellAddress("F4"), new RangeLabel("F4:H8")),
        new PivotTableReport.Source.Range(new SheetName("Budget"), new RangeLabel("A1:C4")),
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

  private static dev.erst.gridgrind.core.model.value.CellReport.TextReport textCell(
      String address, String value) {
    return new dev.erst.gridgrind.core.model.value.CellReport.TextReport(
        new CellAddress(address),
        java.util.Optional.of(value),
        java.util.Optional.of(defaultStyle()),
        java.util.Optional.empty(),
        java.util.Optional.empty(),
        java.util.Optional.of(value),
        java.util.Optional.empty());
  }

  private static CellColorReport rgb(String rgb) {
    return CellColorReport.rgb(new RgbHex(rgb));
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
