package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdSnapshot;
import dev.erst.gridgrind.excel.ExcelDifferentialStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType;
import java.math.BigDecimal;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for protocol-facing conditional-formatting factual report conversion. */
class ConditionalFormattingEntryReportTest {
  @Test
  void convertsConditionalFormattingSnapshotsToProtocolReports() {
    ConditionalFormattingEntryReport entry =
        InspectionResultValidationReportSupport.toConditionalFormattingEntryReport(
            new ExcelConditionalFormattingBlockSnapshot(
                List.of("A1:A3"),
                List.of(
                    new ExcelConditionalFormattingRuleSnapshot.FormulaRule(
                        1,
                        true,
                        "A1>10",
                        new ExcelDifferentialStyleSnapshot(
                            "0.00",
                            true,
                            false,
                            ExcelFontHeight.fromPoints(BigDecimal.valueOf(12)),
                            "#102030",
                            true,
                            true,
                            "#E0F0AA",
                            null,
                            List.of())),
                    new ExcelConditionalFormattingRuleSnapshot.CellValueRule(
                        2, false, ExcelComparisonOperator.BETWEEN, "1", "9", null),
                    new ExcelConditionalFormattingRuleSnapshot.ColorScaleRule(
                        3,
                        false,
                        List.of(
                            new ExcelConditionalFormattingThresholdSnapshot(
                                ExcelConditionalFormattingThresholdType.MIN, null, null),
                            new ExcelConditionalFormattingThresholdSnapshot(
                                ExcelConditionalFormattingThresholdType.MAX, null, null)),
                        List.of("#AA0000", "#00AA00")),
                    new ExcelConditionalFormattingRuleSnapshot.UnsupportedRule(
                        4, false, "FORMULA", "Formula rule is missing formula text."))));

    assertEquals(List.of("A1:A3"), entry.ranges());
    assertEquals(4, entry.rules().size());
    assertInstanceOf(ConditionalFormattingRuleReport.FormulaRule.class, entry.rules().get(0));
    assertInstanceOf(ConditionalFormattingRuleReport.CellValueRule.class, entry.rules().get(1));
    assertInstanceOf(ConditionalFormattingRuleReport.ColorScaleRule.class, entry.rules().get(2));
    ConditionalFormattingRuleReport.UnsupportedRule unsupported =
        assertInstanceOf(
            ConditionalFormattingRuleReport.UnsupportedRule.class, entry.rules().get(3));
    assertEquals("FORMULA", unsupported.kind());
    assertEquals("Formula rule is missing formula text.", unsupported.detail());
  }

  @Test
  void validatesConditionalFormattingHealthReport() {
    ConditionalFormattingHealthReport report =
        new ConditionalFormattingHealthReport(
            3,
            new GridGrindResponse.AnalysisSummaryReport(1, 0, 1, 0),
            List.of(
                new GridGrindResponse.AnalysisFindingReport(
                    AnalysisFindingCode.CONDITIONAL_FORMATTING_PRIORITY_COLLISION,
                    AnalysisSeverity.WARNING,
                    "Priority collision",
                    "Conditional-formatting priorities collide.",
                    new GridGrindResponse.AnalysisLocationReport.Sheet("Ops"),
                    List.of("FORMULA_RULE@Ops!A1:A3"))));

    assertEquals(3, report.checkedConditionalFormattingBlockCount());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingEntryReport(
                List.of(" "),
                List.of(
                    new ConditionalFormattingRuleReport.UnsupportedRule(
                        1, false, "FORMULA", "Missing formula text"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingHealthReport(
                -1, new GridGrindResponse.AnalysisSummaryReport(0, 0, 0, 0), List.of()));
  }
}
