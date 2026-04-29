package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockDefinition;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import dev.erst.gridgrind.excel.ExcelDifferentialBorder;
import dev.erst.gridgrind.excel.ExcelDifferentialBorderSide;
import dev.erst.gridgrind.excel.ExcelDifferentialStyle;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import java.math.BigDecimal;
import java.util.List;
import java.util.Optional;
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
                        Optional.of("#102030"),
                        true,
                        true,
                        Optional.of("#E0F0AA"),
                        Optional.of(
                            new DifferentialBorderInput(
                                Optional.of(
                                    new DifferentialBorderSideInput(
                                        ExcelBorderStyle.THIN, Optional.of("#405060"))),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty(),
                                Optional.empty())))),
                new ConditionalFormattingRuleInput.CellValueRule(
                    ExcelComparisonOperator.BETWEEN,
                    "1",
                    "9",
                    false,
                    new DifferentialStyleInput(
                        null,
                        null,
                        true,
                        null,
                        Optional.empty(),
                        null,
                        null,
                        Optional.of("#AAEECC"),
                        Optional.empty()))));

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
        WorkbookCommandConverter.toExcelConditionalFormattingBlock(input));
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
                            null,
                            true,
                            null,
                            null,
                            Optional.empty(),
                            null,
                            null,
                            Optional.empty(),
                            Optional.empty())))));
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
                            null,
                            true,
                            null,
                            null,
                            Optional.empty(),
                            null,
                            null,
                            Optional.empty(),
                            Optional.empty())))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.FormulaRule(
                " ",
                false,
                new DifferentialStyleInput(
                    null,
                    true,
                    null,
                    null,
                    Optional.empty(),
                    null,
                    null,
                    Optional.empty(),
                    Optional.empty())));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.CellValueRule(
                ExcelComparisonOperator.GREATER_THAN,
                " ",
                null,
                false,
                new DifferentialStyleInput(
                    null,
                    true,
                    null,
                    null,
                    Optional.empty(),
                    null,
                    null,
                    Optional.empty(),
                    Optional.empty())));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DifferentialStyleInput(
                null,
                null,
                null,
                null,
                Optional.empty(),
                null,
                null,
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DifferentialStyleInput(
                " ",
                null,
                null,
                null,
                Optional.empty(),
                null,
                null,
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DifferentialBorderInput(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        NullPointerException.class,
        () -> new DifferentialBorderSideInput(null, Optional.of("#102030")));
  }

  @Test
  void delegatesComparisonRuleSemanticsToEngineNormalization() {
    DifferentialStyleInput style =
        new DifferentialStyleInput(
            null,
            true,
            null,
            null,
            Optional.empty(),
            null,
            null,
            Optional.empty(),
            Optional.empty());

    ConditionalFormattingRuleInput.CellValueRule missingUpperBound =
        new ConditionalFormattingRuleInput.CellValueRule(
            ExcelComparisonOperator.BETWEEN, "1", null, false, style);
    ConditionalFormattingRuleInput.CellValueRule unexpectedUpperBound =
        new ConditionalFormattingRuleInput.CellValueRule(
            ExcelComparisonOperator.GREATER_THAN, "1", "9", false, style);

    IllegalArgumentException missingUpperBoundFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> WorkbookCommandConverter.toExcelConditionalFormattingRule(missingUpperBound));
    IllegalArgumentException unexpectedUpperBoundFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> WorkbookCommandConverter.toExcelConditionalFormattingRule(unexpectedUpperBound));

    assertTrue(missingUpperBoundFailure.getMessage().contains("formula2"));
    assertTrue(unexpectedUpperBoundFailure.getMessage().contains("formula2"));
  }

  @Test
  void convertsDifferentialBordersWithExplicitSides() {
    DifferentialBorderInput border =
        new DifferentialBorderInput(
            Optional.of(
                new DifferentialBorderSideInput(ExcelBorderStyle.THIN, Optional.of("#102030"))),
            Optional.of(
                new DifferentialBorderSideInput(ExcelBorderStyle.DASHED, Optional.of("#203040"))),
            Optional.of(
                new DifferentialBorderSideInput(ExcelBorderStyle.DOUBLE, Optional.of("#304050"))),
            Optional.of(
                new DifferentialBorderSideInput(ExcelBorderStyle.HAIR, Optional.of("#405060"))),
            Optional.of(
                new DifferentialBorderSideInput(ExcelBorderStyle.DOTTED, Optional.of("#506070"))));

    assertEquals(
        new ExcelDifferentialBorder(
            new ExcelDifferentialBorderSide(ExcelBorderStyle.THIN, "#102030"),
            new ExcelDifferentialBorderSide(ExcelBorderStyle.DASHED, "#203040"),
            new ExcelDifferentialBorderSide(ExcelBorderStyle.DOUBLE, "#304050"),
            new ExcelDifferentialBorderSide(ExcelBorderStyle.HAIR, "#405060"),
            new ExcelDifferentialBorderSide(ExcelBorderStyle.DOTTED, "#506070")),
        WorkbookCommandConverter.toExcelDifferentialBorder(border).orElseThrow());
  }

  @Test
  void allowsStyleLessDifferentialRuleFamiliesWhenPoiSupportsThem() {
    assertEquals(
        new ExcelConditionalFormattingRule.FormulaRule("A1>0", false, null),
        WorkbookCommandConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleInput.FormulaRule("A1>0", false, null)));
    assertEquals(
        new ExcelConditionalFormattingRule.CellValueRule(
            ExcelComparisonOperator.GREATER_THAN, "1", null, false, null),
        WorkbookCommandConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleInput.CellValueRule(
                ExcelComparisonOperator.GREATER_THAN, "1", null, false, null)));
    assertEquals(
        new ExcelConditionalFormattingRule.Top10Rule(10, false, false, false, null),
        WorkbookCommandConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleInput.Top10Rule(false, 10, false, false, null)));
  }
}
