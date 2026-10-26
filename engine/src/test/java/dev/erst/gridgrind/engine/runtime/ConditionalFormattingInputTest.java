package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.DifferentialBorderPatch;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.style.FontHeightReport;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelColor;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelDifferentialStyle;
import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for protocol-facing conditional-formatting authoring inputs and conversion. */
class ConditionalFormattingInputTest {
  @Test
  void validatesAndConvertsConditionalFormattingDefinition() {
    ConditionalFormattingSpec input =
        new ConditionalFormattingSpec(
            List.of(
                new ConditionalFormattingRuleSpec.FormulaRule(
                    "A1>0",
                    true,
                    Optional.of(
                        new DifferentialStylePatch(
                            Optional.of("0.00"),
                            Optional.of(true),
                            Optional.of(false),
                            Optional.ofNullable(new FontHeightSpec.Points(BigDecimal.valueOf(11))),
                            Optional.of(ColorSpec.rgb(new RgbHex("#102030"))),
                            Optional.of(true),
                            Optional.of(true),
                            Optional.of(ColorSpec.rgb(new RgbHex("#E0F0AA"))),
                            Optional.of(
                                new DifferentialBorderPatch(
                                    Optional.of(
                                        new BorderSidePatch(
                                            ExcelBorderStyle.THIN,
                                            ColorSpec.rgb(new RgbHex("#405060")))),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty()))))),
                new ConditionalFormattingRuleSpec.CellValueRule(
                    ExcelComparisonOperator.BETWEEN,
                    "1",
                    Optional.of("9"),
                    false,
                    Optional.of(
                        new DifferentialStylePatch(
                            Optional.empty(),
                            Optional.empty(),
                            Optional.of(true),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.of(ColorSpec.rgb(new RgbHex("#AAEECC"))),
                            Optional.empty())))));

    assertEquals(
        new ExcelConditionalFormattingBlockDefinition(
            List.of("A1:A3"),
            List.of(
                new ExcelConditionalFormattingRule.FormulaRule(
                    "A1>0",
                    true,
                    Optional.of(
                        new ExcelDifferentialStyle(
                            Optional.of("0.00"),
                            Optional.of(true),
                            Optional.of(false),
                            Optional.ofNullable(
                                FontHeightReport.fromPoints(BigDecimal.valueOf(11))),
                            Optional.of(ExcelColor.rgb("#102030")),
                            Optional.of(true),
                            Optional.of(true),
                            Optional.of(ExcelColor.rgb("#E0F0AA")),
                            Optional.ofNullable(
                                new ExcelDifferentialBorder(
                                    new ExcelBorderSide(
                                        Optional.of(ExcelBorderStyle.THIN),
                                        Optional.of(ExcelColor.rgb("#405060"))),
                                    null,
                                    null,
                                    null,
                                    null))))),
                new ExcelConditionalFormattingRule.CellValueRule(
                    ExcelComparisonOperator.BETWEEN,
                    "1",
                    Optional.of("9"),
                    false,
                    Optional.of(
                        new ExcelDifferentialStyle(
                            Optional.empty(),
                            Optional.empty(),
                            Optional.of(true),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.of(ExcelColor.rgb("#AAEECC")),
                            Optional.empty()))))),
        WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingBlock(
            List.of("A1:A3"), input));
  }

  @Test
  void rejectsInvalidConditionalFormattingDefinitions() {
    assertThrows(IllegalArgumentException.class, () -> new ConditionalFormattingSpec(List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new ConditionalFormattingSpec(java.util.Collections.singletonList(null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleSpec.FormulaRule(
                " ",
                false,
                Optional.of(
                    new DifferentialStylePatch(
                        Optional.empty(),
                        Optional.of(true),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty()))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleSpec.CellValueRule(
                ExcelComparisonOperator.GREATER_THAN,
                " ",
                Optional.empty(),
                false,
                Optional.of(
                    new DifferentialStylePatch(
                        Optional.empty(),
                        Optional.of(true),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty()))));
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
            new DifferentialBorderPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new BorderSidePatch(Optional.empty(), Optional.empty()));
  }

  @Test
  void delegatesComparisonRuleSemanticsToEngineNormalization() {
    DifferentialStylePatch style =
        new DifferentialStylePatch(
            Optional.empty(),
            Optional.of(true),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());

    IllegalArgumentException missingUpperBoundFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new ConditionalFormattingRuleSpec.CellValueRule(
                    ExcelComparisonOperator.BETWEEN,
                    "1",
                    Optional.empty(),
                    false,
                    Optional.of(style)));
    IllegalArgumentException unexpectedUpperBoundFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new ConditionalFormattingRuleSpec.CellValueRule(
                    ExcelComparisonOperator.GREATER_THAN,
                    "1",
                    Optional.of("9"),
                    false,
                    Optional.of(style)));

    assertTrue(missingUpperBoundFailure.getMessage().contains("formula2"));
    assertTrue(unexpectedUpperBoundFailure.getMessage().contains("formula2"));
  }

  @Test
  void convertsDifferentialBordersWithExplicitSides() {
    DifferentialBorderPatch border =
        new DifferentialBorderPatch(
            Optional.of(
                new BorderSidePatch(ExcelBorderStyle.THIN, ColorSpec.rgb(new RgbHex("#102030")))),
            Optional.of(
                new BorderSidePatch(ExcelBorderStyle.DASHED, ColorSpec.rgb(new RgbHex("#203040")))),
            Optional.of(
                new BorderSidePatch(ExcelBorderStyle.DOUBLE, ColorSpec.rgb(new RgbHex("#304050")))),
            Optional.of(
                new BorderSidePatch(ExcelBorderStyle.HAIR, ColorSpec.rgb(new RgbHex("#405060")))),
            Optional.of(
                new BorderSidePatch(
                    ExcelBorderStyle.DOTTED, ColorSpec.rgb(new RgbHex("#506070")))));

    assertEquals(
        new ExcelDifferentialBorder(
            new ExcelBorderSide(
                Optional.of(ExcelBorderStyle.THIN), Optional.of(ExcelColor.rgb("#102030"))),
            new ExcelBorderSide(
                Optional.of(ExcelBorderStyle.DASHED), Optional.of(ExcelColor.rgb("#203040"))),
            new ExcelBorderSide(
                Optional.of(ExcelBorderStyle.DOUBLE), Optional.of(ExcelColor.rgb("#304050"))),
            new ExcelBorderSide(
                Optional.of(ExcelBorderStyle.HAIR), Optional.of(ExcelColor.rgb("#405060"))),
            new ExcelBorderSide(
                Optional.of(ExcelBorderStyle.DOTTED), Optional.of(ExcelColor.rgb("#506070")))),
        WorkbookCommandStructuredInputConverter.toExcelDifferentialBorder(border).orElseThrow());
  }

  @Test
  void allowsStyleLessDifferentialRuleFamiliesOnlyAsStopIfTrueBarriers() {
    assertEquals(
        new ExcelConditionalFormattingRule.FormulaRule("A1>0", true, Optional.empty()),
        WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleSpec.FormulaRule("A1>0", true, Optional.empty())));
    assertEquals(
        new ExcelConditionalFormattingRule.CellValueRule(
            ExcelComparisonOperator.GREATER_THAN, "1", Optional.empty(), true, Optional.empty()),
        WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleSpec.CellValueRule(
                ExcelComparisonOperator.GREATER_THAN,
                "1",
                Optional.empty(),
                true,
                Optional.empty())));
    assertEquals(
        new ExcelConditionalFormattingRule.Top10Rule(10, false, false, true, Optional.empty()),
        WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleSpec.Top10Rule(true, 10, false, false, Optional.empty())));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingRuleSpec.FormulaRule("A1>0", false, Optional.empty()));
  }
}
