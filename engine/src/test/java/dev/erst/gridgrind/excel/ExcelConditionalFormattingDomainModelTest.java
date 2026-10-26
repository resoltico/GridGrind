package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.DifferentialBorderPatch;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingUnsupportedFeature;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;
import org.apache.poi.ss.usermodel.IconMultiStateFormatting;
import org.junit.jupiter.api.Test;

/** Tests for conditional-formatting engine-side value objects and enum mappings. */
class ExcelConditionalFormattingDomainModelTest {
  @Test
  void mapsEverySupportedPoiIconSetAndThresholdType() {
    assertEquals(
        ExcelConditionalFormattingIconSet.GYR_3_ARROW,
        ExcelConditionalFormattingPoiBridge.fromPoi(IconMultiStateFormatting.IconSet.GYR_3_ARROW));
    assertEquals(
        ExcelConditionalFormattingIconSet.GREY_3_ARROWS,
        ExcelConditionalFormattingPoiBridge.fromPoi(
            IconMultiStateFormatting.IconSet.GREY_3_ARROWS));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYR_3_FLAGS,
        ExcelConditionalFormattingPoiBridge.fromPoi(IconMultiStateFormatting.IconSet.GYR_3_FLAGS));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS,
        ExcelConditionalFormattingPoiBridge.fromPoi(
            IconMultiStateFormatting.IconSet.GYR_3_TRAFFIC_LIGHTS));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYR_3_TRAFFIC_LIGHTS_BOX,
        ExcelConditionalFormattingPoiBridge.fromPoi(
            IconMultiStateFormatting.IconSet.GYR_3_TRAFFIC_LIGHTS_BOX));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYR_3_SHAPES,
        ExcelConditionalFormattingPoiBridge.fromPoi(IconMultiStateFormatting.IconSet.GYR_3_SHAPES));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYR_3_SYMBOLS_CIRCLE,
        ExcelConditionalFormattingPoiBridge.fromPoi(
            IconMultiStateFormatting.IconSet.GYR_3_SYMBOLS_CIRCLE));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYR_3_SYMBOLS,
        ExcelConditionalFormattingPoiBridge.fromPoi(
            IconMultiStateFormatting.IconSet.GYR_3_SYMBOLS));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYR_4_ARROWS,
        ExcelConditionalFormattingPoiBridge.fromPoi(IconMultiStateFormatting.IconSet.GYR_4_ARROWS));
    assertEquals(
        ExcelConditionalFormattingIconSet.GREY_4_ARROWS,
        ExcelConditionalFormattingPoiBridge.fromPoi(
            IconMultiStateFormatting.IconSet.GREY_4_ARROWS));
    assertEquals(
        ExcelConditionalFormattingIconSet.RB_4_TRAFFIC_LIGHTS,
        ExcelConditionalFormattingPoiBridge.fromPoi(
            IconMultiStateFormatting.IconSet.RB_4_TRAFFIC_LIGHTS));
    assertEquals(
        ExcelConditionalFormattingIconSet.RATINGS_4,
        ExcelConditionalFormattingPoiBridge.fromPoi(IconMultiStateFormatting.IconSet.RATINGS_4));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYRB_4_TRAFFIC_LIGHTS,
        ExcelConditionalFormattingPoiBridge.fromPoi(
            IconMultiStateFormatting.IconSet.GYRB_4_TRAFFIC_LIGHTS));
    assertEquals(
        ExcelConditionalFormattingIconSet.GYYYR_5_ARROWS,
        ExcelConditionalFormattingPoiBridge.fromPoi(
            IconMultiStateFormatting.IconSet.GYYYR_5_ARROWS));
    assertEquals(
        ExcelConditionalFormattingIconSet.GREY_5_ARROWS,
        ExcelConditionalFormattingPoiBridge.fromPoi(
            IconMultiStateFormatting.IconSet.GREY_5_ARROWS));
    assertEquals(
        ExcelConditionalFormattingIconSet.RATINGS_5,
        ExcelConditionalFormattingPoiBridge.fromPoi(IconMultiStateFormatting.IconSet.RATINGS_5));
    assertEquals(
        ExcelConditionalFormattingIconSet.QUARTERS_5,
        ExcelConditionalFormattingPoiBridge.fromPoi(IconMultiStateFormatting.IconSet.QUARTERS_5));

    assertEquals(
        ExcelConditionalFormattingThresholdType.NUMBER,
        ExcelConditionalFormattingPoiBridge.fromPoi(
            ConditionalFormattingThreshold.RangeType.NUMBER));
    assertEquals(
        ExcelConditionalFormattingThresholdType.MIN,
        ExcelConditionalFormattingPoiBridge.fromPoi(ConditionalFormattingThreshold.RangeType.MIN));
    assertEquals(
        ExcelConditionalFormattingThresholdType.MAX,
        ExcelConditionalFormattingPoiBridge.fromPoi(ConditionalFormattingThreshold.RangeType.MAX));
    assertEquals(
        ExcelConditionalFormattingThresholdType.PERCENT,
        ExcelConditionalFormattingPoiBridge.fromPoi(
            ConditionalFormattingThreshold.RangeType.PERCENT));
    assertEquals(
        ExcelConditionalFormattingThresholdType.PERCENTILE,
        ExcelConditionalFormattingPoiBridge.fromPoi(
            ConditionalFormattingThreshold.RangeType.PERCENTILE));
    assertEquals(
        ExcelConditionalFormattingThresholdType.UNALLOCATED,
        ExcelConditionalFormattingPoiBridge.fromPoi(
            ConditionalFormattingThreshold.RangeType.UNALLOCATED));
    assertEquals(
        ExcelConditionalFormattingThresholdType.FORMULA,
        ExcelConditionalFormattingPoiBridge.fromPoi(
            ConditionalFormattingThreshold.RangeType.FORMULA));
  }

  @Test
  void enforcesDifferentialStyleAndBorderInvariants() {
    BorderSidePatch side =
        new BorderSidePatch(
            Optional.of(ExcelBorderStyle.THIN), Optional.of(ColorSpec.rgb(new RgbHex("#102030"))));
    DifferentialBorderPatch border =
        new DifferentialBorderPatch(
            Optional.of(side),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    DifferentialStylePatch style =
        new DifferentialStylePatch(
            Optional.of("0.00"),
            Optional.of(true),
            Optional.of(false),
            Optional.ofNullable(
                new dev.erst.gridgrind.core.model.style.FontHeightSpec.Points(
                    BigDecimal.valueOf(11))),
            Optional.of(new ColorSpec.Theme(3, Optional.of(0.25d))),
            Optional.of(true),
            Optional.of(false),
            Optional.of(new ColorSpec.Indexed(12, Optional.empty())),
            Optional.ofNullable(border));
    List<ExcelConditionalFormattingUnsupportedFeature> unsupportedFeatures =
        new ArrayList<>(List.of(ExcelConditionalFormattingUnsupportedFeature.FONT_ATTRIBUTES));
    ExcelDifferentialStyleSnapshot snapshot =
        new ExcelDifferentialStyleSnapshot(
            "0.00",
            true,
            false,
            FontHeightReport.fromPoints(BigDecimal.valueOf(11)),
            new ColorSpec.Theme(3, Optional.of(0.25d)),
            true,
            false,
            new ColorSpec.Indexed(12, Optional.empty()),
            border,
            unsupportedFeatures);
    unsupportedFeatures.clear();

    assertEquals(Optional.of(new ColorSpec.Theme(3, Optional.of(0.25d))), style.fontColor());
    assertEquals(Optional.of(new ColorSpec.Indexed(12, Optional.empty())), style.fillColor());
    assertEquals(
        List.of(ExcelConditionalFormattingUnsupportedFeature.FONT_ATTRIBUTES),
        snapshot.unsupportedFeatures());

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DifferentialBorderPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertDoesNotThrow(
        () ->
            new DifferentialBorderPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(side)));
    assertDoesNotThrow(
        () ->
            new DifferentialStylePatch(
                Optional.empty(),
                Optional.of(true),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertDoesNotThrow(
        () ->
            new DifferentialStylePatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.ofNullable(border)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DifferentialStylePatch(
                Optional.of(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DifferentialStylePatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
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
            Optional.of(
                new DifferentialStylePatch(
                    Optional.of("0.00"),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())));
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
    DifferentialStylePatch style =
        new DifferentialStylePatch(
            Optional.of("0.00"),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    ExcelDifferentialStyleSnapshot snapshot =
        new ExcelDifferentialStyleSnapshot(
            "0.00", null, null, null, null, null, null, null, null, List.of());
    ExcelConditionalFormattingThresholdSnapshot threshold =
        new ExcelConditionalFormattingThresholdSnapshot(
            ExcelConditionalFormattingThresholdType.MIN, null, 0.0d);

    ExcelConditionalFormattingRule.CellValueRule authoredCellValueRule =
        new ExcelConditionalFormattingRule.CellValueRule(
            ExcelComparisonOperator.BETWEEN, "=1", Optional.of("=9"), false, Optional.of(style));
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
    assertEquals("9", authoredCellValueRule.formula2().orElseThrow());
    assertEquals(List.of("#102030"), colorScaleRule.colors());
    assertEquals(ExcelConditionalFormattingIconSet.GYR_3_ARROW, iconSetRule.iconSet());

    assertDoesNotThrow(
        () -> new ExcelConditionalFormattingRule.FormulaRule("=A1>0", true, Optional.empty()));
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
                1, false, "#102030", false, -1, 100, threshold, threshold));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelConditionalFormattingRuleSnapshot.DataBarRule(
                1, false, "#102030", false, 0, -1, threshold, threshold));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelConditionalFormattingRuleSnapshot.DataBarRule(
                1, false, "#102030", false, 0, 100, null, threshold));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelConditionalFormattingRuleSnapshot.IconSetRule(
                1, false, null, false, false, List.of(threshold)));
  }
}
