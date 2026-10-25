package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.chart.ChartReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.protection.CellProtectionReport;
import dev.erst.gridgrind.core.model.protection.SheetProtectionReport;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentReport;
import dev.erst.gridgrind.core.model.style.CellBorderReport;
import dev.erst.gridgrind.core.model.style.CellBorderSideReport;
import dev.erst.gridgrind.core.model.style.CellFillReport;
import dev.erst.gridgrind.core.model.style.CellFontReport;
import dev.erst.gridgrind.core.model.style.CellStyleReport;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.model.value.CellScalarValue;
import dev.erst.gridgrind.core.result.SheetSummaryReport;
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
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Exhaustive assertion-type coverage for the canonical Phase-4 verification contract. */
class AssertionCoverageTest {
  @Test
  void assertionResultsExposeTaggedEvidenceAndRejectMismatchedIdentity() {
    AssertionFailure failure =
        new AssertionFailure(
            new StepId("assert-cell"),
            "EXPECT_CELL_VALUE",
            new dev.erst.gridgrind.core.plan.CellSelector.ByAddress(
                new SheetName("Budget"), new CellAddress("A1")),
            new CellAssertion.CellValue(new CellScalarValue.Text("expected")),
            List.of());
    AssertionResult.Passed passed =
        new AssertionResult.Passed(new StepId("assert-pass"), "EXPECT_CELL_VALUE");
    AssertionResult.Failed failed =
        new AssertionResult.Failed(new StepId("assert-cell"), "EXPECT_CELL_VALUE", failure);
    assertEquals(AssertionOutcome.PASSED, passed.outcome());
    assertEquals(AssertionOutcome.FAILED, failed.outcome());
    assertEquals(failure, failed.failure());
    assertThrows(
        IllegalArgumentException.class,
        () -> new AssertionResult.Failed(new StepId("other"), "EXPECT_CELL_VALUE", failure));
    assertThrows(
        IllegalArgumentException.class,
        () -> new AssertionResult.Failed(new StepId("assert-cell"), "OTHER", failure));
  }

  @Test
  void exposesStableDiscriminatorsAcrossEveryAssertionLeaf() {
    CellStyleReport style = style();
    NamedRangeReport namedRange =
        new NamedRangeReport.RangeReport(
            "BudgetTotal",
            new NamedRangeScope.Workbook(),
            "Budget!B2",
            dev.erst.gridgrind.core.model.formula.NamedRangeTarget.range(
                new SheetName("Budget"), new RangeRef("B2")));

    assertEquals("EXPECT_SHEET_PRESENT", new PresenceAssertion.SheetPresent().assertionType());
    assertEquals("EXPECT_SHEET_ABSENT", new PresenceAssertion.SheetAbsent().assertionType());
    assertEquals(
        "EXPECT_NAMED_RANGE_PRESENT", new PresenceAssertion.NamedRangePresent().assertionType());
    assertEquals(
        "EXPECT_NAMED_RANGE_ABSENT", new PresenceAssertion.NamedRangeAbsent().assertionType());
    assertEquals("EXPECT_TABLE_PRESENT", new PresenceAssertion.TablePresent().assertionType());
    assertEquals("EXPECT_TABLE_ABSENT", new PresenceAssertion.TableAbsent().assertionType());
    assertEquals(
        "EXPECT_PIVOT_TABLE_PRESENT", new PresenceAssertion.PivotTablePresent().assertionType());
    assertEquals(
        "EXPECT_PIVOT_TABLE_ABSENT", new PresenceAssertion.PivotTableAbsent().assertionType());
    assertEquals("EXPECT_CHART_PRESENT", new PresenceAssertion.ChartPresent().assertionType());
    assertEquals("EXPECT_CHART_ABSENT", new PresenceAssertion.ChartAbsent().assertionType());
    assertEquals(
        "EXPECT_CELL_VALUE",
        new CellAssertion.CellValue(
                new dev.erst.gridgrind.core.model.value.CellScalarValue.Text("Owner"))
            .assertionType());
    assertEquals("EXPECT_DISPLAY_VALUE", new CellAssertion.DisplayValue("Owner").assertionType());
    assertEquals("EXPECT_FORMULA_TEXT", new CellAssertion.FormulaText("2+3").assertionType());
    assertEquals("EXPECT_CELL_STYLE", new CellAssertion.CellStyle(style).assertionType());
    assertEquals(
        "EXPECT_WORKBOOK_PROTECTION",
        new WorkbookFactAssertion.WorkbookProtectionFacts(
                new WorkbookProtectionReport(true, false, false, false, false))
            .assertionType());
    assertEquals(
        "EXPECT_SHEET_STRUCTURE",
        new WorkbookFactAssertion.SheetStructureFacts(
                new SheetSummaryReport(
                    new SheetLabel("Budget"),
                    ExcelSheetVisibility.VISIBLE,
                    new SheetProtectionReport.Unprotected(),
                    2,
                    1,
                    1))
            .assertionType());
    assertEquals(
        "EXPECT_NAMED_RANGE_FACTS",
        new WorkbookFactAssertion.NamedRangeFacts(List.of(namedRange)).assertionType());
    assertEquals(
        "EXPECT_TABLE_FACTS",
        new WorkbookFactAssertion.TableFacts(List.of(tableReport())).assertionType());
    assertEquals(
        "EXPECT_PIVOT_TABLE_FACTS",
        new WorkbookFactAssertion.PivotTableFacts(List.of(pivotReport())).assertionType());
    assertEquals(
        "EXPECT_CHART_FACTS",
        new WorkbookFactAssertion.ChartFacts(List.of(chartReport())).assertionType());
    assertEquals(
        "EXPECT_ANALYSIS_MAX_SEVERITY",
        new AnalysisAssertion.AnalysisMaxSeverity(
                new InspectionAnalysisQuery.AnalyzeFormulaHealth(), AnalysisSeverity.WARNING)
            .assertionType());
    assertEquals(
        "EXPECT_ANALYSIS_FINDING_PRESENT",
        new AnalysisAssertion.AnalysisFindingPresent(
                new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                AnalysisFindingCode.FORMULA_ERROR_RESULT,
                Optional.of(AnalysisSeverity.ERROR),
                Optional.of("error"))
            .assertionType());
    assertEquals(
        "EXPECT_ANALYSIS_FINDING_ABSENT",
        new AnalysisAssertion.AnalysisFindingAbsent(
                new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                Optional.empty(),
                Optional.empty())
            .assertionType());
    assertEquals(
        "ALL_OF",
        new CompositeAssertion.AllOf(
                List.of(
                    new CellAssertion.CellValue(
                        new dev.erst.gridgrind.core.model.value.CellScalarValue.Text("Owner")),
                    new CellAssertion.DisplayValue("Owner")))
            .assertionType());
    assertEquals(
        "ANY_OF",
        new CompositeAssertion.AnyOf(
                List.of(
                    new CellAssertion.CellValue(
                        new dev.erst.gridgrind.core.model.value.CellScalarValue.Text("Owner")),
                    new CellAssertion.DisplayValue("Owner")))
            .assertionType());
    assertEquals(
        "NOT",
        new CompositeAssertion.Not(
                new CellAssertion.CellValue(
                    new dev.erst.gridgrind.core.model.value.CellScalarValue.Text("Owner")))
            .assertionType());
  }

  @Test
  void expectedCellValueVariantsCoverSuccessConstructors() {
    assertInstanceOf(CellScalarValue.Blank.class, new CellScalarValue.Blank());
    assertTrue(new CellScalarValue.BooleanValue(true).bool());
    assertEquals("#REF!", new CellScalarValue.ErrorValue("#REF!").error());
    assertEquals(42.0d, new CellScalarValue.NumberValue(42.0d).number());
  }

  @Test
  void assertionSupportCoversCollectionAndAnalysisValidationBranches() {
    NamedRangeReport namedRange =
        new NamedRangeReport.FormulaReport(
            "BudgetExpr", new NamedRangeScope.Workbook(), "SUM(Budget!B2:B4)");
    InspectionResult observation =
        new WorkbookInspectionResult.WorkbookSummaryResult(
            new StepId("summary"),
            new WorkbookSummary.WithSheets(
                1,
                List.of(new SheetLabel("Budget")),
                new SheetLabel("Budget"),
                List.of(new SheetLabel("Budget")),
                1,
                false));

    assertEquals(
        List.of(new PresenceAssertion.TablePresent()),
        AssertionSupport.copyAssertions(
            List.of(new PresenceAssertion.TablePresent()), "assertions"));
    assertEquals(
        List.of(namedRange), AssertionSupport.copyNamedRanges(List.of(namedRange), "namedRanges"));
    assertEquals(
        List.of(tableReport()), AssertionSupport.copyTables(List.of(tableReport()), "tables"));
    assertEquals(
        List.of(pivotReport()),
        AssertionSupport.copyPivotTables(List.of(pivotReport()), "pivotTables"));
    assertEquals(
        List.of(chartReport()), AssertionSupport.copyCharts(List.of(chartReport()), "charts"));
    assertEquals(
        List.of(observation),
        AssertionSupport.copyObservations(List.of(observation), "observations"));
    InspectionQuery.Analysis analysis =
        AssertionSupport.requireAnalysisQuery(
            new InspectionAnalysisQuery.AnalyzeFormulaHealth(), "query");
    assertInstanceOf(InspectionAnalysisQuery.AnalyzeFormulaHealth.class, analysis);

    assertEquals(
        "query must be an analysis query",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    AssertionSupport.requireAnalysisQuery(
                        new SheetIntrospectionQuery.GetCells(), "query"))
            .getMessage());
    assertEquals(
        "assertions must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> AssertionSupport.copyAssertions(List.of(), "assertions"))
            .getMessage());
    assertEquals(
        "tables must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    AssertionSupport.copyTables(
                        java.util.Arrays.asList(tableReport(), null), "tables"))
            .getMessage());
    assertEquals(
        "charts must not be null",
        assertThrows(NullPointerException.class, () -> AssertionSupport.copyCharts(null, "charts"))
            .getMessage());
  }

  @Test
  void analysisFindingConstructorsCoverOptionalMessageBranches() {
    AnalysisAssertion.AnalysisFindingPresent findingPresent =
        new AnalysisAssertion.AnalysisFindingPresent(
            new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
            AnalysisFindingCode.FORMULA_ERROR_RESULT,
            Optional.of(AnalysisSeverity.ERROR),
            Optional.empty());
    AnalysisAssertion.AnalysisFindingAbsent findingAbsent =
        new AnalysisAssertion.AnalysisFindingAbsent(
            new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
            AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
            Optional.of(AnalysisSeverity.INFO),
            Optional.of("volatile"));

    assertEquals(Optional.empty(), findingPresent.messageContains());
    assertEquals(Optional.of("volatile"), findingAbsent.messageContains());
    assertEquals(
        "messageContains must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new AnalysisAssertion.AnalysisFindingPresent(
                        new InspectionAnalysisQuery.AnalyzeFormulaHealth(),
                        AnalysisFindingCode.FORMULA_ERROR_RESULT,
                        Optional.of(AnalysisSeverity.ERROR),
                        Optional.of(" ")))
            .getMessage());
  }

  private static CellStyleReport style() {
    CellBorderSideReport emptySide = new CellBorderSideReport.None();
    return new CellStyleReport(
        "General",
        new CellAlignmentReport(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new CellFontReport(
            false,
            false,
            "Aptos",
            new FontHeightReport(220, BigDecimal.valueOf(11)),
            null,
            false,
            false),
        CellFillReport.pattern(ExcelFillPattern.NONE),
        new CellBorderReport(emptySide, emptySide, emptySide, emptySide),
        new CellProtectionReport(true, false));
  }

  private static TableEntryReport tableReport() {
    return new TableEntryReport(
        "BudgetTable",
        new SheetLabel("Budget"),
        new RangeLabel("A1:B2"),
        1,
        0,
        List.of("Owner", "Amount"),
        new TableStyleReport.None(),
        true);
  }

  private static PivotTableReport pivotReport() {
    return new PivotTableReport.Supported(
        "Sales Pivot",
        new SheetName("Report"),
        new PivotTableReport.Anchor(new CellAddress("C5"), new RangeLabel("C5:G9")),
        new PivotTableReport.Source.Range(new SheetName("Data"), new RangeLabel("A1:D5")),
        List.of(new PivotTableReport.Field(0, "Region")),
        List.of(new PivotTableReport.Field(1, "Stage")),
        List.of(new PivotTableReport.Field(2, "Owner")),
        List.of(
            new PivotTableReport.DataField(
                3,
                "Amount",
                ExcelPivotDataConsolidateFunction.SUM,
                "Total Amount",
                Optional.of("#,##0.00"))),
        true);
  }

  private static ChartReport chartReport() {
    return new ChartReport(
        "Revenue",
        new dev.erst.gridgrind.core.model.drawing.DrawingAnchorReport.Absolute(
            1L, 2L, 3L, 4L, ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
        new ChartReport.Title.Text("Revenue"),
        new ChartReport.Legend.Visible(ExcelChartLegendPosition.RIGHT),
        ExcelChartDisplayBlanksAs.GAP,
        true,
        List.of(
            new ChartReport.Bar(
                false,
                ExcelChartBarDirection.COLUMN,
                ExcelChartBarGrouping.CLUSTERED,
                Optional.empty(),
                Optional.empty(),
                List.of(
                    new ChartReport.Axis(
                        ExcelChartAxisKind.CATEGORY,
                        ExcelChartAxisPosition.BOTTOM,
                        ExcelChartAxisCrosses.AUTO_ZERO,
                        true)),
                List.of(
                    new ChartReport.Series(
                        new ChartReport.Title.Text("Series 1"),
                        new ChartReport.DataSource.StringLiteral(List.of("Jan", "Feb")),
                        new ChartReport.DataSource.NumericLiteral(
                            Optional.of("#,##0"), List.of("1", "2")),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty())))));
  }
}
