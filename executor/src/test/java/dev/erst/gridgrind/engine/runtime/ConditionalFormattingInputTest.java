package dev.erst.gridgrind.engine.runtime;

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
  void validatesAndConvertsConditionalFormattingDefinition() {
    ConditionalFormattingDefinitionInput input =
        new ConditionalFormattingDefinitionInput(
            List.of(
                new ConditionalFormattingRuleInput.FormulaRule(
                    "A1>0",
                    true,
                    Optional.of(
                        new DifferentialStyleInput(
                            Optional.of("0.00"),
                            Optional.of(true),
                            Optional.of(false),
                            Optional.ofNullable(new FontHeightInput.Points(BigDecimal.valueOf(11))),
                            Optional.of("#102030"),
                            Optional.of(true),
                            Optional.of(true),
                            Optional.of("#E0F0AA"),
                            Optional.of(
                                new DifferentialBorderInput(
                                    Optional.of(
                                        new DifferentialBorderSideInput(
                                            ExcelBorderStyle.THIN, Optional.of("#405060"))),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty(),
                                    Optional.empty()))))),
                new ConditionalFormattingRuleInput.CellValueRule(
                    ExcelComparisonOperator.BETWEEN,
                    "1",
                    Optional.of("9"),
                    false,
                    Optional.of(
                        new DifferentialStyleInput(
                            Optional.empty(),
                            Optional.empty(),
                            Optional.of(true),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.of("#AAEECC"),
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
                            Optional.ofNullable(ExcelFontHeight.fromPoints(BigDecimal.valueOf(11))),
                            Optional.of("#102030"),
                            Optional.of(true),
                            Optional.of(true),
                            Optional.of("#E0F0AA"),
                            Optional.ofNullable(
                                new ExcelDifferentialBorder(
                                    new ExcelDifferentialBorderSide(
                                        ExcelBorderStyle.THIN, "#405060"),
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
                            Optional.of("#AAEECC"),
                            Optional.empty()))))),
        WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingBlock(
            List.of("A1:A3"), input));
  }

  @Test
  void rejectsInvalidConditionalFormattingDefinitions() {
    assertThrows(
        IllegalArgumentException.class, () -> new ConditionalFormattingDefinitionInput(List.of()));
    assertThrows(
        NullPointerException.class,
        () -> new ConditionalFormattingDefinitionInput(java.util.Collections.singletonList(null)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.FormulaRule(
                " ",
                false,
                Optional.of(
                    new DifferentialStyleInput(
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
            new ConditionalFormattingRuleInput.CellValueRule(
                ExcelComparisonOperator.GREATER_THAN,
                " ",
                Optional.empty(),
                false,
                Optional.of(
                    new DifferentialStyleInput(
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
            new DifferentialStyleInput(
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
            new DifferentialStyleInput(
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
                new ConditionalFormattingRuleInput.CellValueRule(
                    ExcelComparisonOperator.BETWEEN,
                    "1",
                    Optional.empty(),
                    false,
                    Optional.of(style)));
    IllegalArgumentException unexpectedUpperBoundFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new ConditionalFormattingRuleInput.CellValueRule(
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
        WorkbookCommandStructuredInputConverter.toExcelDifferentialBorder(border).orElseThrow());
  }

  @Test
  void allowsStyleLessDifferentialRuleFamiliesWhenPoiSupportsThem() {
    assertEquals(
        new ExcelConditionalFormattingRule.FormulaRule("A1>0", false, Optional.empty()),
        WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleInput.FormulaRule("A1>0", false, Optional.empty())));
    assertEquals(
        new ExcelConditionalFormattingRule.CellValueRule(
            ExcelComparisonOperator.GREATER_THAN, "1", Optional.empty(), false, Optional.empty()),
        WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleInput.CellValueRule(
                ExcelComparisonOperator.GREATER_THAN,
                "1",
                Optional.empty(),
                false,
                Optional.empty())));
    assertEquals(
        new ExcelConditionalFormattingRule.Top10Rule(10, false, false, false, Optional.empty()),
        WorkbookCommandStructuredInputConverter.toExcelConditionalFormattingRule(
            new ConditionalFormattingRuleInput.Top10Rule(
                false, 10, false, false, Optional.empty())));
  }
}
