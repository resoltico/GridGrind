package dev.erst.gridgrind.contract.assertion;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.dto.AnalysisFindingCode;
import dev.erst.gridgrind.contract.dto.AnalysisSeverity;
import dev.erst.gridgrind.contract.dto.CellAlignmentReport;
import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.TableStyleReport;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisCrosses;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisKind;
import dev.erst.gridgrind.excel.foundation.ExcelChartAxisPosition;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Exhaustive assertion-type coverage for the canonical Phase-4 verification contract. */
class AssertionCoverageTest {
  @Test
  void exposesStableDiscriminatorsAcrossEveryAssertionLeaf() {
    GridGrindResponse.CellStyleReport style = style();
    GridGrindResponse.NamedRangeReport namedRange =
        new GridGrindResponse.NamedRangeReport.RangeReport(
            "BudgetTotal",
            new NamedRangeScope.Workbook(),
            "Budget!B2",
            new dev.erst.gridgrind.contract.dto.NamedRangeTarget("Budget", "B2"));

    assertEquals("EXPECT_NAMED_RANGE_PRESENT", new Assertion.NamedRangePresent().assertionType());
    assertEquals("EXPECT_NAMED_RANGE_ABSENT", new Assertion.NamedRangeAbsent().assertionType());
    assertEquals("EXPECT_TABLE_PRESENT", new Assertion.TablePresent().assertionType());
    assertEquals("EXPECT_TABLE_ABSENT", new Assertion.TableAbsent().assertionType());
    assertEquals("EXPECT_PIVOT_TABLE_PRESENT", new Assertion.PivotTablePresent().assertionType());
    assertEquals("EXPECT_PIVOT_TABLE_ABSENT", new Assertion.PivotTableAbsent().assertionType());
    assertEquals("EXPECT_CHART_PRESENT", new Assertion.ChartPresent().assertionType());
    assertEquals("EXPECT_CHART_ABSENT", new Assertion.ChartAbsent().assertionType());
    assertEquals(
        "EXPECT_CELL_VALUE",
        new Assertion.CellValue(new ExpectedCellValue.Text("Owner")).assertionType());
    assertEquals("EXPECT_DISPLAY_VALUE", new Assertion.DisplayValue("Owner").assertionType());
    assertEquals("EXPECT_FORMULA_TEXT", new Assertion.FormulaText("2+3").assertionType());
    assertEquals("EXPECT_CELL_STYLE", new Assertion.CellStyle(style).assertionType());
    assertEquals(
        "EXPECT_WORKBOOK_PROTECTION",
        new Assertion.WorkbookProtectionFacts(
                new WorkbookProtectionReport(true, false, false, false, false))
            .assertionType());
    assertEquals(
        "EXPECT_SHEET_STRUCTURE",
        new Assertion.SheetStructureFacts(
                new GridGrindResponse.SheetSummaryReport(
                    "Budget",
                    ExcelSheetVisibility.VISIBLE,
                    new GridGrindResponse.SheetProtectionReport.Unprotected(),
                    2,
                    1,
                    1))
            .assertionType());
    assertEquals(
        "EXPECT_NAMED_RANGE_FACTS",
        new Assertion.NamedRangeFacts(List.of(namedRange)).assertionType());
    assertEquals(
        "EXPECT_TABLE_FACTS", new Assertion.TableFacts(List.of(tableReport())).assertionType());
    assertEquals(
        "EXPECT_PIVOT_TABLE_FACTS",
        new Assertion.PivotTableFacts(List.of(pivotReport())).assertionType());
    assertEquals(
        "EXPECT_CHART_FACTS", new Assertion.ChartFacts(List.of(chartReport())).assertionType());
    assertEquals(
        "EXPECT_ANALYSIS_MAX_SEVERITY",
        new Assertion.AnalysisMaxSeverity(
                new InspectionQuery.AnalyzeFormulaHealth(), AnalysisSeverity.WARNING)
            .assertionType());
    assertEquals(
        "EXPECT_ANALYSIS_FINDING_PRESENT",
        new Assertion.AnalysisFindingPresent(
                new InspectionQuery.AnalyzeFormulaHealth(),
                AnalysisFindingCode.FORMULA_ERROR_RESULT,
                AnalysisSeverity.ERROR,
                "error")
            .assertionType());
    assertEquals(
        "EXPECT_ANALYSIS_FINDING_ABSENT",
        new Assertion.AnalysisFindingAbsent(
                new InspectionQuery.AnalyzeFormulaHealth(),
                AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                null,
                null)
            .assertionType());
    assertEquals(
        "ALL_OF",
        new Assertion.AllOf(
                List.of(
                    new Assertion.CellValue(new ExpectedCellValue.Text("Owner")),
                    new Assertion.DisplayValue("Owner")))
            .assertionType());
    assertEquals(
        "ANY_OF",
        new Assertion.AnyOf(
                List.of(
                    new Assertion.CellValue(new ExpectedCellValue.Text("Owner")),
                    new Assertion.DisplayValue("Owner")))
            .assertionType());
    assertEquals(
        "NOT",
        new Assertion.Not(new Assertion.CellValue(new ExpectedCellValue.Text("Owner")))
            .assertionType());
  }

  @Test
  void expectedCellValueVariantsCoverSuccessConstructors() {
    assertInstanceOf(ExpectedCellValue.Blank.class, new ExpectedCellValue.Blank());
    assertEquals(true, new ExpectedCellValue.BooleanValue(true).value());
    assertEquals("#REF!", new ExpectedCellValue.ErrorValue("#REF!").error());
    assertEquals(42.0d, new ExpectedCellValue.NumericValue(42.0d).number());
  }

  @Test
  void assertionSupportCoversCollectionAndAnalysisValidationBranches() {
    GridGrindResponse.NamedRangeReport namedRange =
        new GridGrindResponse.NamedRangeReport.FormulaReport(
            "BudgetExpr", new NamedRangeScope.Workbook(), "SUM(Budget!B2:B4)");
    InspectionResult observation =
        new InspectionResult.WorkbookSummaryResult(
            "summary",
            new GridGrindResponse.WorkbookSummary.WithSheets(
                1, List.of("Budget"), "Budget", List.of("Budget"), 1, false));

    assertEquals(
        List.of(new Assertion.TablePresent()),
        AssertionSupport.copyAssertions(List.of(new Assertion.TablePresent()), "assertions"));
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
        AssertionSupport.requireAnalysisQuery(new InspectionQuery.AnalyzeFormulaHealth(), "query");
    assertInstanceOf(InspectionQuery.AnalyzeFormulaHealth.class, analysis);

    assertEquals(
        "query must be an analysis query",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    AssertionSupport.requireAnalysisQuery(new InspectionQuery.GetCells(), "query"))
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
    Assertion.AnalysisFindingPresent findingPresent =
        new Assertion.AnalysisFindingPresent(
            new InspectionQuery.AnalyzeFormulaHealth(),
            AnalysisFindingCode.FORMULA_ERROR_RESULT,
            AnalysisSeverity.ERROR,
            null);
    Assertion.AnalysisFindingAbsent findingAbsent =
        new Assertion.AnalysisFindingAbsent(
            new InspectionQuery.AnalyzeFormulaHealth(),
            AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
            AnalysisSeverity.INFO,
            "volatile");

    assertNull(findingPresent.messageContains());
    assertEquals("volatile", findingAbsent.messageContains());
    assertEquals(
        "messageContains must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new Assertion.AnalysisFindingPresent(
                        new InspectionQuery.AnalyzeFormulaHealth(),
                        AnalysisFindingCode.FORMULA_ERROR_RESULT,
                        AnalysisSeverity.ERROR,
                        " "))
            .getMessage());
  }

  private static GridGrindResponse.CellStyleReport style() {
    CellBorderSideReport emptySide = new CellBorderSideReport(ExcelBorderStyle.NONE, null);
    return new GridGrindResponse.CellStyleReport(
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
        new CellFillReport(ExcelFillPattern.NONE, null, null),
        new CellBorderReport(emptySide, emptySide, emptySide, emptySide),
        new CellProtectionReport(true, false));
  }

  private static TableEntryReport tableReport() {
    return new TableEntryReport(
        "BudgetTable",
        "Budget",
        "A1:B2",
        1,
        0,
        List.of("Owner", "Amount"),
        new TableStyleReport.None(),
        true);
  }

  private static PivotTableReport pivotReport() {
    return new PivotTableReport.Supported(
        "Sales Pivot",
        "Report",
        new PivotTableReport.Anchor("C5", "C5:G9"),
        new PivotTableReport.Source.Range("Data", "A1:D5"),
        List.of(new PivotTableReport.Field(0, "Region")),
        List.of(new PivotTableReport.Field(1, "Stage")),
        List.of(new PivotTableReport.Field(2, "Owner")),
        List.of(
            new PivotTableReport.DataField(
                3, "Amount", ExcelPivotDataConsolidateFunction.SUM, "Total Amount", "#,##0.00")),
        true);
  }

  private static ChartReport chartReport() {
    return new ChartReport(
        "Revenue",
        new dev.erst.gridgrind.contract.dto.DrawingAnchorReport.Absolute(
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
                null,
                null,
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
                        new ChartReport.DataSource.NumericLiteral("#,##0", List.of("1", "2")),
                        null,
                        null,
                        null,
                        null)))));
  }
}
