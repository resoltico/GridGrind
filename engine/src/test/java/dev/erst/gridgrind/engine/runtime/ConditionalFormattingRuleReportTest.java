package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelColor;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdSnapshot;
import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelDifferentialStyleSnapshot;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingUnsupportedFeature;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for conditional-formatting protocol report shapes and conversions. */
class ConditionalFormattingRuleReportTest {
  @Test
  void convertsAdvancedConditionalFormattingRuleFamiliesFromExcel() {
    ExcelConditionalFormattingThresholdSnapshot minThreshold =
        new ExcelConditionalFormattingThresholdSnapshot(
            ExcelConditionalFormattingThresholdType.MIN, null, 0.0d);
    ExcelConditionalFormattingThresholdSnapshot maxThreshold =
        new ExcelConditionalFormattingThresholdSnapshot(
            ExcelConditionalFormattingThresholdType.MAX, null, 100.0d);

    ConditionalFormattingRuleReport dataBar =
        InspectionResultValidationReportSupport.toConditionalFormattingRuleReport(
            new ExcelConditionalFormattingRuleSnapshot.DataBarRule(
                4, false, "#102030", true, 10, 90, minThreshold, maxThreshold));
    ConditionalFormattingRuleReport iconSet =
        InspectionResultValidationReportSupport.toConditionalFormattingRuleReport(
            new ExcelConditionalFormattingRuleSnapshot.IconSetRule(
                5,
                true,
                ExcelConditionalFormattingIconSet.GYR_3_ARROW,
                false,
                true,
                List.of(minThreshold, maxThreshold)));

    ConditionalFormattingRuleReport.DataBarRule dataBarRule =
        assertInstanceOf(ConditionalFormattingRuleReport.DataBarRule.class, dataBar);
    ConditionalFormattingRuleReport.IconSetRule iconSetRule =
        assertInstanceOf(ConditionalFormattingRuleReport.IconSetRule.class, iconSet);

    assertEquals("#102030", dataBarRule.color());
    assertEquals(10, dataBarRule.widthMin());
    assertEquals(90, dataBarRule.widthMax());
    assertEquals(ExcelConditionalFormattingIconSet.GYR_3_ARROW, iconSetRule.iconSet());
    assertEquals(2, iconSetRule.thresholds().size());
  }

  @Test
  void validatesConditionalFormattingRuleReportInvariants() {
    ConditionalFormattingThresholdReport threshold =
        new ConditionalFormattingThresholdReport(
            ExcelConditionalFormattingThresholdType.NUMBER, null, 5.0d);

    assertEquals(
        "1",
        new ConditionalFormattingRuleReport.CellValueRule(
                1,
                false,
                ExcelComparisonOperator.GREATER_THAN,
                "1",
                Optional.empty(),
                Optional.empty())
            .formula1());
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
                1, false, "#102030", false, -1, 90, threshold, threshold));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleReport.DataBarRule(
                1, false, "#102030", false, 10, -1, threshold, threshold));
    assertThrows(
        NullPointerException.class,
        () ->
            new ConditionalFormattingRuleReport.IconSetRule(
                1, false, null, false, false, List.of(threshold)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleReport.ColorScaleRule(
                -1, false, List.of(threshold), List.of("#102030")));
  }

  @Test
  void convertsAndValidatesDifferentialStyleAndBorderReports() {
    ExcelDifferentialBorder border =
        new ExcelDifferentialBorder(
            borderSide(ExcelBorderStyle.THIN, "#102030"),
            borderSide(ExcelBorderStyle.DASHED, "#203040"),
            borderSide(ExcelBorderStyle.DOUBLE, "#304050"),
            borderSide(ExcelBorderStyle.HAIR, "#405060"),
            borderSide(ExcelBorderStyle.DOTTED, "#506070"));
    ExcelDifferentialStyleSnapshot style =
        new ExcelDifferentialStyleSnapshot(
            null,
            true,
            false,
            null,
            ExcelColor.rgb("#111111"),
            true,
            false,
            ExcelColor.rgb("#EEEEEE"),
            border,
            List.of(ExcelConditionalFormattingUnsupportedFeature.ALIGNMENT));

    DifferentialStyleReport report =
        InspectionResultValidationReportSupport.toDifferentialStyleReport(style).orElseThrow();
    DifferentialBorderReport borderReport =
        InspectionResultValidationReportSupport.toDifferentialBorderReport(border).orElseThrow();
    DifferentialBorderReport sparseBorderReport =
        InspectionResultValidationReportSupport.toDifferentialBorderReport(
                new ExcelDifferentialBorder(
                    borderSide(ExcelBorderStyle.THIN, "#102030"), null, null, null, null))
            .orElseThrow();
    DifferentialBorderSideReport borderSideReport =
        InspectionResultValidationReportSupport.toDifferentialBorderSideReport(
                borderSide(ExcelBorderStyle.THICK, "#AABBCC"))
            .orElseThrow();

    assertTrue(InspectionResultValidationReportSupport.toDifferentialStyleReport(null).isEmpty());
    assertTrue(InspectionResultValidationReportSupport.toDifferentialBorderReport(null).isEmpty());
    assertEquals(Optional.of(CellColorReport.rgb("#111111")), report.fontColor());
    assertEquals(Optional.of(CellColorReport.rgb("#EEEEEE")), report.fillColor());
    assertEquals(
        List.of(ExcelConditionalFormattingUnsupportedFeature.ALIGNMENT),
        report.unsupportedFeatures());
    assertEquals(Optional.of(ExcelBorderStyle.DASHED), borderReport.top().orElseThrow().style());
    assertTrue(sparseBorderReport.top().isEmpty());
    assertEquals(Optional.of(CellColorReport.rgb("#AABBCC")), borderSideReport.color());

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
    assertTrue(
        InspectionResultValidationReportSupport.toDifferentialBorderSideReport(null).isEmpty());
  }

  @Test
  void convertsAndValidatesThresholdReports() {
    ConditionalFormattingThresholdReport threshold =
        InspectionResultValidationReportSupport.toConditionalFormattingThresholdReport(
            new ExcelConditionalFormattingThresholdSnapshot(
                ExcelConditionalFormattingThresholdType.FORMULA, "A1", null));

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

  private static ExcelBorderSide borderSide(ExcelBorderStyle style, String rgb) {
    return new ExcelBorderSide(Optional.of(style), Optional.of(ExcelColor.rgb(rgb)));
  }
}
