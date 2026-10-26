package dev.erst.gridgrind.core.model.validation;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.style.DifferentialBorderReport;
import dev.erst.gridgrind.core.model.style.DifferentialStyleReport;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingUnsupportedFeature;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for conditional-formatting protocol report shapes and conversions. */
class ConditionalFormattingRuleReportTest {
  @Test
  void validatesConditionalFormattingRuleReportInvariants() {
    var threshold =
        new ConditionalFormattingThresholdReport(
            ExcelConditionalFormattingThresholdType.NUMBER, null, 1.0d);
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingRuleReport.FormulaRule(1, false, " ", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleReport.CellValueRule(
                1,
                false,
                ExcelComparisonOperator.BETWEEN,
                " ",
                Optional.of("9"),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleReport.CellValueRule(
                1,
                false,
                ExcelComparisonOperator.BETWEEN,
                "1",
                Optional.of(" "),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingRuleReport.UnsupportedRule(1, false, " ", "detail"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingRuleReport.UnsupportedRule(1, false, "FORMULA", " "));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleReport.DataBarRule(
                1, false, Optional.of(new RgbHex("#102030")), false, -1, 90, threshold, threshold));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleReport.DataBarRule(
                1, false, Optional.of(new RgbHex("#102030")), false, 10, -1, threshold, threshold));
    assertThrows(
        NullPointerException.class,
        () ->
            new ConditionalFormattingRuleReport.IconSetRule(
                1, false, null, false, false, List.of(threshold)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleReport.ColorScaleRule(
                -1, false, List.of(threshold), List.of(new RgbHex("#102030"))));
  }

  @Test
  void validatesDifferentialStyleAndBorderReports() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DifferentialStyleReport(
                " ",
                null,
                null,
                null,
                Optional.empty(),
                null,
                null,
                Optional.empty(),
                Optional.empty(),
                List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DifferentialStyleReport(
                null,
                null,
                null,
                null,
                Optional.empty(),
                null,
                null,
                Optional.empty(),
                Optional.empty(),
                List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new DifferentialStyleReport(
                "0.00",
                null,
                null,
                null,
                Optional.empty(),
                null,
                null,
                Optional.empty(),
                Optional.empty(),
                List.of((ExcelConditionalFormattingUnsupportedFeature) null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DifferentialBorderReport(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
  }

  @Test
  void validatesThresholdReports() {
    var threshold =
        new ConditionalFormattingThresholdReport(
            ExcelConditionalFormattingThresholdType.FORMULA, "A1", null);

    assertEquals(ExcelConditionalFormattingThresholdType.FORMULA, threshold.type());
    assertEquals("A1", threshold.formula());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingThresholdReport(
                ExcelConditionalFormattingThresholdType.FORMULA, " ", null));
    assertThrows(
        NullPointerException.class,
        () -> new ConditionalFormattingThresholdReport(null, null, null));
  }
}
