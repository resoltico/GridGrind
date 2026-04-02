package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelDifferentialBorderSide;
import dev.erst.gridgrind.excel.ExcelDifferentialStyle;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import java.math.BigDecimal;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for protocol-facing conditional-formatting authoring inputs and conversion. */
class ConditionalFormattingInputTest {
  @Test
  void validatesAndConvertsConditionalFormattingBlock() {
    ConditionalFormattingBlockInput input =
        new ConditionalFormattingBlockInput(
            List.of("A1:A3"),
            List.of(
                new ConditionalFormattingRuleInput.FormulaRule(
                    "A1>0",
                    true,
                    new DifferentialStyleInput(
                        "0.00",
                        true,
                        false,
                        new FontHeightInput.Points(BigDecimal.valueOf(11)),
                        "#102030",
                        true,
                        true,
                        "#E0F0AA",
                        new DifferentialBorderInput(
                            new DifferentialBorderSideInput(ExcelBorderStyle.THIN, "#405060"),
                            null,
                            null,
                            null,
                            null))),
                new ConditionalFormattingRuleInput.CellValueRule(
                    ExcelComparisonOperator.BETWEEN,
                    "1",
                    "9",
                    false,
                    new DifferentialStyleInput(
                        null, null, true, null, null, null, null, "#AAEECC", null))));

    assertEquals(
        new ExcelConditionalFormattingBlockDefinition(
            List.of("A1:A3"),
            List.of(
                new ExcelConditionalFormattingRule.FormulaRule(
                    "A1>0",
                    true,
                    new ExcelDifferentialStyle(
                        "0.00",
                        true,
                        false,
                        ExcelFontHeight.fromPoints(BigDecimal.valueOf(11)),
                        "#102030",
                        true,
                        true,
                        "#E0F0AA",
                        new ExcelDifferentialBorder(
                            new ExcelDifferentialBorderSide(ExcelBorderStyle.THIN, "#405060"),
                            null,
                            null,
                            null,
                            null))),
                new ExcelConditionalFormattingRule.CellValueRule(
                    ExcelComparisonOperator.BETWEEN,
                    "1",
                    "9",
                    false,
                    new ExcelDifferentialStyle(
                        null, null, true, null, null, null, null, "#AAEECC", null)))),
        input.toExcelConditionalFormattingBlockDefinition());
  }

  @Test
  void rejectsInvalidConditionalFormattingInputs() {
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingBlockInput(List.of(), List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ConditionalFormattingBlockInput(List.of("A1:A3"), List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingBlockInput(
                List.of(" "),
                List.of(
                    new ConditionalFormattingRuleInput.FormulaRule(
                        "A1>0",
                        false,
                        new DifferentialStyleInput(
                            null, true, null, null, null, null, null, null, null)))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingBlockInput(
                List.of("A1:A3", "A1:A3"),
                List.of(
                    new ConditionalFormattingRuleInput.FormulaRule(
                        "A1>0",
                        false,
                        new DifferentialStyleInput(
                            null, true, null, null, null, null, null, null, null)))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.FormulaRule(
                " ",
                false,
                new DifferentialStyleInput(null, true, null, null, null, null, null, null, null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.CellValueRule(
                ExcelComparisonOperator.GREATER_THAN,
                " ",
                null,
                false,
                new DifferentialStyleInput(null, true, null, null, null, null, null, null, null)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DifferentialStyleInput(null, null, null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DifferentialStyleInput(" ", null, null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DifferentialBorderInput(null, null, null, null, null));
    assertThrows(
        NullPointerException.class, () -> new DifferentialBorderSideInput(null, "#102030"));
  }

  @Test
  void delegatesComparisonRuleSemanticsToEngineNormalization() {
    DifferentialStyleInput style =
        new DifferentialStyleInput(null, true, null, null, null, null, null, null, null);

    ConditionalFormattingRuleInput.CellValueRule missingUpperBound =
        new ConditionalFormattingRuleInput.CellValueRule(
            ExcelComparisonOperator.BETWEEN, "1", null, false, style);
    ConditionalFormattingRuleInput.CellValueRule unexpectedUpperBound =
        new ConditionalFormattingRuleInput.CellValueRule(
            ExcelComparisonOperator.GREATER_THAN, "1", "9", false, style);

    IllegalArgumentException missingUpperBoundFailure =
        assertThrows(
            IllegalArgumentException.class, missingUpperBound::toExcelConditionalFormattingRule);
    IllegalArgumentException unexpectedUpperBoundFailure =
        assertThrows(
            IllegalArgumentException.class, unexpectedUpperBound::toExcelConditionalFormattingRule);

    assertTrue(missingUpperBoundFailure.getMessage().contains("formula2"));
    assertTrue(unexpectedUpperBoundFailure.getMessage().contains("formula2"));
  }

  @Test
  void convertsDifferentialBordersWithExplicitSides() {
    DifferentialBorderInput border =
        new DifferentialBorderInput(
            new DifferentialBorderSideInput(ExcelBorderStyle.THIN, "#102030"),
            new DifferentialBorderSideInput(ExcelBorderStyle.DASHED, "#203040"),
            new DifferentialBorderSideInput(ExcelBorderStyle.DOUBLE, "#304050"),
            new DifferentialBorderSideInput(ExcelBorderStyle.HAIR, "#405060"),
            new DifferentialBorderSideInput(ExcelBorderStyle.DOTTED, "#506070"));

    assertEquals(
        new ExcelDifferentialBorder(
            new ExcelDifferentialBorderSide(ExcelBorderStyle.THIN, "#102030"),
            new ExcelDifferentialBorderSide(ExcelBorderStyle.DASHED, "#203040"),
            new ExcelDifferentialBorderSide(ExcelBorderStyle.DOUBLE, "#304050"),
            new ExcelDifferentialBorderSide(ExcelBorderStyle.HAIR, "#405060"),
            new ExcelDifferentialBorderSide(ExcelBorderStyle.DOTTED, "#506070")),
        border.toExcelDifferentialBorder());
  }
}
