package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting;
import org.junit.jupiter.api.Test;

/** Tests for conditional-formatting engine-side value objects and enum mappings. */
class ExcelConditionalFormattingDomainModelTest {
  @Test
  void mapsEverySupportedPoiIconSetAndThresholdType() {
    assertEquals(
        ExcelConditionalFormattingIconSet.GYR_3_ARROW,
        ExcelConditionalFormattingIconSet.fromPoi(IconMultiStateFormatting.IconSet.GYR_3_ARROW));
    assertEquals(
        ExcelConditionalFormattingIconSet.GREY_3_ARROWS,
        ExcelConditionalFormattingIconSet.fromPoi(IconMultiStateFormatting.IconSet.GREY_3_ARROWS));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYR_3_FLAGS,
        ExcelConditionalFormattingIconSet.fromPoi(IconMultiStateFormatting.IconSet.GYR_3_FLAGS));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
        ExcelConditionalFormattingIconSet.fromPoi(
            IconMultiStateFormatting.IconSet.GYR_3_TRAFFIC_LIGHTS));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS_BOX,
        ExcelConditionalFormattingIconSet.fromPoi(
            IconMultiStateFormatting.IconSet.GYR_3_TRAFFIC_LIGHTS_BOX));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYR_3_SHAPES,
        ExcelConditionalFormattingIconSet.fromPoi(IconMultiStateFormatting.IconSet.GYR_3_SHAPES));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYR_3_SYMBOLS_CIRCLE,
        ExcelConditionalFormattingIconSet.fromPoi(
            IconMultiStateFormatting.IconSet.GYR_3_SYMBOLS_CIRCLE));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYR_3_SYMBOLS,
        ExcelConditionalFormattingIconSet.fromPoi(IconMultiStateFormatting.IconSet.GYR_3_SYMBOLS));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYR_4_ARROWS,
        ExcelConditionalFormattingIconSet.fromPoi(IconMultiStateFormatting.IconSet.GYR_4_ARROWS));
    assertEquals(
        ExcelConditionalFormattingIconSet.GREY_4_ARROWS,
        ExcelConditionalFormattingIconSet.fromPoi(IconMultiStateFormatting.IconSet.GREY_4_ARROWS));
    assertEquals(
        ExcelConditionalFormattingIconSet.RB_4_TRAFFIC_LIGHTS,
        ExcelConditionalFormattingIconSet.fromPoi(
            IconMultiStateFormatting.IconSet.RB_4_TRAFFIC_LIGHTS));
    assertEquals(
        ExcelConditionalFormattingIconSet.RATINGS_4,
        ExcelConditionalFormattingIconSet.fromPoi(IconMultiStateFormatting.IconSet.RATINGS_4));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYRB_4_TRAFFIC_LIGHTS,
        ExcelConditionalFormattingIconSet.fromPoi(
            IconMultiStateFormatting.IconSet.GYRB_4_TRAFFIC_LIGHTS));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYYYR_5_ARROWS,
        ExcelConditionalFormattingIconSet.fromPoi(IconMultiStateFormatting.IconSet.GYYYR_5_ARROWS));
    assertEquals(
        ExcelConditionalFormattingIconSet.GREY_5_ARROWS,
        ExcelConditionalFormattingIconSet.fromPoi(IconMultiStateFormatting.IconSet.GREY_5_ARROWS));
    assertEquals(
        ExcelConditionalFormattingIconSet.RATINGS_5,
        ExcelConditionalFormattingIconSet.fromPoi(IconMultiStateFormatting.IconSet.RATINGS_5));
    assertEquals(
        ExcelConditionalFormattingIconSet.QUARTERS_5,
        ExcelConditionalFormattingIconSet.fromPoi(IconMultiStateFormatting.IconSet.QUARTERS_5));

    assertEquals(
        ExcelConditionalFormattingThresholdType.NUMBER,
        ExcelConditionalFormattingThresholdType.fromPoi(
            ConditionalFormattingThreshold.RangeType.NUMBER));
    assertEquals(
        ExcelConditionalFormattingThresholdType.MIN,
        ExcelConditionalFormattingThresholdType.fromPoi(
            ConditionalFormattingThreshold.RangeType.MIN));
    assertEquals(
        ExcelConditionalFormattingThresholdType.MAX,
        ExcelConditionalFormattingThresholdType.fromPoi(
            ConditionalFormattingThreshold.RangeType.MAX));
    assertEquals(
        ExcelConditionalFormattingThresholdType.PERCENT,
        ExcelConditionalFormattingThresholdType.fromPoi(
            ConditionalFormattingThreshold.RangeType.PERCENT));
    assertEquals(
        ExcelConditionalFormattingThresholdType.PERCENTILE,
        ExcelConditionalFormattingThresholdType.fromPoi(
            ConditionalFormattingThreshold.RangeType.PERCENTILE));
    assertEquals(
        ExcelConditionalFormattingThresholdType.UNALLOCATED,
        ExcelConditionalFormattingThresholdType.fromPoi(
            ConditionalFormattingThreshold.RangeType.UNALLOCATED));
    assertEquals(
        ExcelConditionalFormattingThresholdType.FORMULA,
        ExcelConditionalFormattingThresholdType.fromPoi(
            ConditionalFormattingThreshold.RangeType.FORMULA));
  }

  @Test
  void enforcesDifferentialStyleAndBorderInvariants() {
    ExcelDifferentialBorderSide side =
        new ExcelDifferentialBorderSide(ExcelBorderStyle.THIN, "#102030");
    ExcelDifferentialBorder border = new ExcelDifferentialBorder(side, null, null, null, null);
    ExcelDifferentialStyle style =
        new ExcelDifferentialStyle(
            "0.00",
            true,
            false,
            ExcelFontHeight.fromPoints(BigDecimal.valueOf(11)),
            "#223344",
            true,
            false,
            "#AABBCC",
            border);
    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures =
        new ArrayList<>(List.of(ExcelConditionalFormattingUnsupportedFeature.FONT_ATTRIBUTES));
    ExcelDifferentialStyleSnapshot snapshot =
        new ExcelDifferentialStyleSnapshot(
            "0.00",
            true,
            false,
            ExcelFontHeight.fromPoints(BigDecimal.valueOf(11)),
            "#223344",
            true,
            false,
            "#AABBCC",
            border,
            unsupportedFeatures);
    unsupportedFeatures.clear();

    assertEquals("#223344", style.fontColor());
    assertEquals("#AABBCC", style.fillColor());
    assertEquals(
        List.of(ExcelConditionalFormattingUnsupportedFeature.FONT_ATTRIBUTES),
        snapshot.unsupportedFeatures());

    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelDifferentialBorder(null, null, null, null, null));
    assertDoesNotThrow(() -> new ExcelDifferentialBorder(null, null, null, null, side));
    assertDoesNotThrow(
        () -> new ExcelDifferentialStyle(null, true, null, null, null, null, null, null, null));
    assertDoesNotThrow(
        () -> new ExcelDifferentialStyle(null, null, null, null, null, null, null, null, border));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelDifferentialStyle(" ", null, null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelDifferentialStyle(null, null, null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDifferentialStyleSnapshot(
                " ", null, null, null, null, null, null, null, null, List.of()));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelDifferentialStyleSnapshot(
                null, null, null, null, null, null, null, null, null, null));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelDifferentialStyleSnapshot(
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                border,
                List.of(ExcelConditionalFormattingUnsupportedFeature.ALIGNMENT, null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDifferentialStyleSnapshot(
                null, null, null, null, null, null, null, null, null, List.of()));
    assertDoesNotThrow(
        () ->
            new ExcelDifferentialStyleSnapshot(
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                List.of(ExcelConditionalFormattingUnsupportedFeature.ALIGNMENT)));
  }

  @Test
  void enforcesBlockAndThresholdInvariants() {
    ExcelConditionalFormattingRule.FormulaRule rule =
        new ExcelConditionalFormattingRule.FormulaRule(
            "=A1>0",
            true,
            new ExcelDifferentialStyle("0.00", null, null, null, null, null, null, null, null));
    ExcelConditionalFormattingBlockDefinition definition =
        new ExcelConditionalFormattingBlockDefinition(List.of("A1:A3"), List.of(rule));

    assertEquals(List.of("A1:A3"), definition.ranges());
    assertEquals("A1>0", rule.formula());
    assertDoesNotThrow(
        () ->
            new ExcelConditionalFormattingThresholdSnapshot(
                ExcelConditionalFormattingThresholdType.NUMBER, null, 1.0d));
    assertDoesNotThrow(
        () ->
            new ExcelConditionalFormattingThresholdSnapshot(
                ExcelConditionalFormattingThresholdType.FORMULA, "A1", null));
    assertDoesNotThrow(
        () ->
            new ExcelConditionalFormattingBlockSnapshot(
                List.of("A1:A3"),
                List.of(
                    new ExcelConditionalFormattingRuleSnapshot.UnsupportedRule(
                        1, false, "FORMULA", "detail"))));

    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelConditionalFormattingBlockDefinition(List.of(), List.of(rule)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingBlockDefinition(
                List.of("A1:A3", "A1:A3"), List.of(rule)));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelConditionalFormattingBlockDefinition(List.of("A1:A3"), List.of(rule, null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingBlockSnapshot(
                List.of(" "),
                List.of(
                    new ExcelConditionalFormattingRuleSnapshot.UnsupportedRule(
                        1, false, "FORMULA", "detail"))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelConditionalFormattingBlockDefinition(List.of("A1:A3"), List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelConditionalFormattingThresholdSnapshot(null, null, 1.0d));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingThresholdSnapshot(
                ExcelConditionalFormattingThresholdType.FORMULA, " ", null));
  }

  @Test
  void enforcesAuthoredAndLoadedRuleInvariants() {
    ExcelDifferentialStyle style =
        new ExcelDifferentialStyle("0.00", null, null, null, null, null, null, null, null);
    ExcelDifferentialStyleSnapshot snapshot =
        new ExcelDifferentialStyleSnapshot(
            "0.00", null, null, null, null, null, null, null, null, List.of());
    ExcelConditionalFormattingThresholdSnapshot threshold =
        new ExcelConditionalFormattingThresholdSnapshot(
            ExcelConditionalFormattingThresholdType.MIN, null, 0.0d);

    ExcelConditionalFormattingRule.CellValueRule authoredCellValueRule =
        new ExcelConditionalFormattingRule.CellValueRule(
            ExcelComparisonOperator.BETWEEN, "=1", "=9", false, style);
    ExcelConditionalFormattingRuleSnapshot.ColorScaleRule colorScaleRule =
        new ExcelConditionalFormattingRuleSnapshot.ColorScaleRule(
            1, false, List.of(threshold), List.of("#102030"));
    ExcelConditionalFormattingRuleSnapshot.IconSetRule iconSetRule =
        new ExcelConditionalFormattingRuleSnapshot.IconSetRule(
            2,
            false,
            ExcelConditionalFormattingIconSet.GYR_3_ARROW,
            false,
            false,
            List.of(threshold));

    assertEquals("1", authoredCellValueRule.formula1());
    assertEquals("9", authoredCellValueRule.formula2());
    assertEquals(List.of("#102030"), colorScaleRule.colors());
    assertEquals(ExcelConditionalFormattingIconSet.GYR_3_ARROW, iconSetRule.iconSet());

    assertThrows(
        NullPointerException.class,
        () -> new ExcelConditionalFormattingRule.FormulaRule("=A1>0", true, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelConditionalFormattingRuleSnapshot.FormulaRule(-1, false, "A1>0", snapshot));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRuleSnapshot.CellValueRule(
                1, false, ExcelComparisonOperator.GREATER_THAN, "1", " ", snapshot));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRuleSnapshot.DataBarRule(
                1, false, "#102030", false, true, -1, 100, threshold, threshold));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRuleSnapshot.DataBarRule(
                1, false, "#102030", false, true, 0, -1, threshold, threshold));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelConditionalFormattingRuleSnapshot.DataBarRule(
                1, false, "#102030", false, true, 0, 100, null, threshold));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelConditionalFormattingRuleSnapshot.IconSetRule(
                1, false, null, false, false, List.of(threshold)));
  }
}
