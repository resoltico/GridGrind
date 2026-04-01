package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for workbook-read result invariants. */
class WorkbookReadResultTest {
  @Test
  void workbookSummaryResultRejectsBlankRequestId() {
    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new WorkbookReadResult.WorkbookSummaryResult(
                    " ", new GridGrindResponse.WorkbookSummary(1, List.of("Budget"), 0, false)));

    assertEquals("requestId must not be blank", exception.getMessage());
  }

  @Test
  void cellsResultCopiesCells() {
    WorkbookReadResult.CellsResult result =
        new WorkbookReadResult.CellsResult(
            "cells",
            "Budget",
            List.of(
                new GridGrindResponse.CellReport.BlankReport(
                    "A1", "BLANK", "", defaultStyle(), null, null)));

    assertEquals("A1", result.cells().getFirst().address());
  }

  @Test
  void analysisReadResultsRejectBlankRequestIdAndNullAnalysis() {
    GridGrindResponse.AnalysisSummaryReport summary =
        new GridGrindResponse.AnalysisSummaryReport(1, 0, 0, 1);
    GridGrindResponse.AnalysisFindingReport finding =
        new GridGrindResponse.AnalysisFindingReport(
            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.INFO,
            "Volatile formula",
            "Formula uses NOW().",
            new GridGrindResponse.AnalysisLocationReport.Workbook(),
            List.of("NOW()"));
    GridGrindResponse.HyperlinkHealthReport hyperlinkHealth =
        new GridGrindResponse.HyperlinkHealthReport(1, summary, List.of(finding));
    GridGrindResponse.NamedRangeHealthReport namedRangeHealth =
        new GridGrindResponse.NamedRangeHealthReport(1, summary, List.of(finding));
    GridGrindResponse.WorkbookFindingsReport workbookFindings =
        new GridGrindResponse.WorkbookFindingsReport(summary, List.of(finding));

    assertEquals(
        1,
        new WorkbookReadResult.HyperlinkHealthResult("hyperlink-health", hyperlinkHealth)
            .analysis()
            .checkedHyperlinkCount());
    assertEquals(
        1,
        new WorkbookReadResult.NamedRangeHealthResult("named-range-health", namedRangeHealth)
            .analysis()
            .checkedNamedRangeCount());
    assertEquals(
        1,
        new WorkbookReadResult.WorkbookFindingsResult("workbook-findings", workbookFindings)
            .analysis()
            .summary()
            .totalCount());

    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.HyperlinkHealthResult(" ", hyperlinkHealth));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.NamedRangeHealthResult("named-range-health", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.WorkbookFindingsResult(" ", workbookFindings));
  }

  @Test
  void dataValidationResultsCopyEntriesAndRejectInvalidState() {
    WorkbookReadResult.DataValidationsResult result =
        new WorkbookReadResult.DataValidationsResult(
            "validations",
            "Budget",
            List.of(
                new DataValidationEntryReport.Supported(
                    List.of("A2:A5"),
                    new DataValidationEntryReport.DataValidationDefinitionReport(
                        new DataValidationRuleInput.WholeNumber(
                            ExcelComparisonOperator.GREATER_OR_EQUAL, "1", null),
                        true,
                        false,
                        new DataValidationPromptInput("Priority", "Use 1 or greater.", true),
                        null))));

    assertEquals("A2:A5", result.validations().getFirst().ranges().getFirst());
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.DataValidationsResult(" ", "Budget", List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.DataValidationsResult("validations", "Budget", null));
  }

  @Test
  void dataValidationHealthResultsRejectBlankRequestIdAndNullAnalysis() {
    GridGrindResponse.AnalysisSummaryReport summary =
        new GridGrindResponse.AnalysisSummaryReport(2, 1, 1, 0);
    GridGrindResponse.AnalysisFindingReport finding =
        new GridGrindResponse.AnalysisFindingReport(
            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                .DATA_VALIDATION_OVERLAPPING_RULES,
            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.WARNING,
            "Overlapping data validations",
            "Rules overlap on the same cells.",
            new GridGrindResponse.AnalysisLocationReport.Range("Budget", "A3:A4"),
            List.of("A1:A5", "A3:A4"));
    DataValidationHealthReport health =
        new DataValidationHealthReport(2, summary, List.of(finding));

    assertEquals(
        2,
        new WorkbookReadResult.DataValidationHealthResult("validation-health", health)
            .analysis()
            .checkedValidationCount());
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookReadResult.DataValidationHealthResult(" ", health));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookReadResult.DataValidationHealthResult("validation-health", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationHealthReport(-1, summary, List.of(finding)));
    assertThrows(
        NullPointerException.class,
        () -> new DataValidationHealthReport(1, null, List.of(finding)));
    assertThrows(
        NullPointerException.class,
        () -> new DataValidationHealthReport(1, summary, List.of(finding, null)));
  }

  private static GridGrindResponse.CellStyleReport defaultStyle() {
    return new GridGrindResponse.CellStyleReport(
        "General",
        false,
        false,
        false,
        dev.erst.gridgrind.excel.ExcelHorizontalAlignment.GENERAL,
        dev.erst.gridgrind.excel.ExcelVerticalAlignment.BOTTOM,
        "Aptos",
        new FontHeightReport(220, java.math.BigDecimal.valueOf(11)),
        null,
        false,
        false,
        null,
        dev.erst.gridgrind.excel.ExcelBorderStyle.NONE,
        dev.erst.gridgrind.excel.ExcelBorderStyle.NONE,
        dev.erst.gridgrind.excel.ExcelBorderStyle.NONE,
        dev.erst.gridgrind.excel.ExcelBorderStyle.NONE);
  }
}
