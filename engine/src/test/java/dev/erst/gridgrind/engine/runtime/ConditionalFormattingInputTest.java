package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.DifferentialBorderPatch;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
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
                        new DifferentialStylePatch(
                            Optional.of("0.00"),
                            Optional.of(true),
                            Optional.of(false),
                            Optional.ofNullable(
                                new dev.erst.gridgrind.core.model.style.FontHeightSpec.Points(
                                    BigDecimal.valueOf(11))),
                            Optional.of(ColorSpec.rgb(new RgbHex("#102030"))),
                            Optional.of(true),
                            Optional.of(true),
                            Optional.of(ColorSpec.rgb(new RgbHex("#E0F0AA"))),
                            Optional.ofNullable(
                                new DifferentialBorderPatch(
                                    Optional.of(
                                        new BorderSidePatch(
                                            Optional.of(ExcelBorderStyle.THIN),
                                            Optional.of(ColorSpec.rgb(new RgbHex("#405060"))))),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty()))))),
                new ExcelConditionalFormattingRule.CellValueRule(
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
