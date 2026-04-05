package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingUnsupportedFeature;
import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelDifferentialBorderSide;
import dev.erst.gridgrind.excel.ExcelDifferentialStyleSnapshot;
import dev.erst.gridgrind.protocol.dto.*;
import java.util.List;
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
        WorkbookReadResultConverter.toConditionalFormattingRuleReport(
            new ExcelConditionalFormattingRuleSnapshot.DataBarRule(
                4, false, "#102030", true, false, 10, 90, minThreshold, maxThreshold));
    ConditionalFormattingRuleReport iconSet =
        WorkbookReadResultConverter.toConditionalFormattingRuleReport(
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
                1, false, ExcelComparisonOperator.GREATER_THAN, "1", null, null)
            .formula1());
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingRuleReport.FormulaRule(1, false, " ", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleReport.CellValueRule(
                1, false, ExcelComparisonOperator.BETWEEN, " ", "9", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleReport.CellValueRule(
                1, false, ExcelComparisonOperator.BETWEEN, "1", " ", null));
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
                1, false, "#102030", false, true, -1, 90, threshold, threshold));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleReport.DataBarRule(
                1, false, "#102030", false, true, 10, -1, threshold, threshold));
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
            new ExcelDifferentialBorderSide(ExcelBorderStyle.THIN, "#102030"),
            new ExcelDifferentialBorderSide(ExcelBorderStyle.DASHED, "#203040"),
            new ExcelDifferentialBorderSide(ExcelBorderStyle.DOUBLE, "#304050"),
            new ExcelDifferentialBorderSide(ExcelBorderStyle.HAIR, "#405060"),
            new ExcelDifferentialBorderSide(ExcelBorderStyle.DOTTED, "#506070"));
    ExcelDifferentialStyleSnapshot style =
        new ExcelDifferentialStyleSnapshot(
            null,
            true,
            false,
            null,
            "#111111",
            true,
            false,
            "#EEEEEE",
            border,
            List.of(ExcelConditionalFormattingUnsupportedFeature.ALIGNMENT));

    DifferentialStyleReport report = WorkbookReadResultConverter.toDifferentialStyleReport(style);
    DifferentialBorderReport borderReport =
        WorkbookReadResultConverter.toDifferentialBorderReport(border);
    DifferentialBorderReport sparseBorderReport =
        WorkbookReadResultConverter.toDifferentialBorderReport(
            new ExcelDifferentialBorder(
                new ExcelDifferentialBorderSide(ExcelBorderStyle.THIN, "#102030"),
                null,
                null,
                null,
                null));
    DifferentialBorderSideReport borderSideReport =
        WorkbookReadResultConverter.toDifferentialBorderSideReport(
            new ExcelDifferentialBorderSide(ExcelBorderStyle.THICK, "#AABBCC"));

    assertNull(WorkbookReadResultConverter.toDifferentialStyleReport(null));
    assertNull(WorkbookReadResultConverter.toDifferentialBorderReport(null));
    assertEquals("#111111", report.fontColor());
    assertEquals("#EEEEEE", report.fillColor());
    assertEquals(
        List.of(ExcelConditionalFormattingUnsupportedFeature.ALIGNMENT),
        report.unsupportedFeatures());
    assertEquals(ExcelBorderStyle.DASHED, borderReport.top().style());
    assertNull(sparseBorderReport.top());
    assertEquals("#AABBCC", borderSideReport.color());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DifferentialStyleReport(
                " ", null, null, null, null, null, null, null, null, List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DifferentialStyleReport(
                null, null, null, null, null, null, null, null, null, List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new DifferentialStyleReport(
                "0.00",
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                List.of((ExcelConditionalFormattingUnsupportedFeature) null)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DifferentialBorderReport(null, null, null, null, null));
    assertNull(WorkbookReadResultConverter.toDifferentialBorderSideReport(null));
  }

  @Test
  void convertsAndValidatesThresholdReports() {
    ConditionalFormattingThresholdReport threshold =
        WorkbookReadResultConverter.toConditionalFormattingThresholdReport(
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
}
