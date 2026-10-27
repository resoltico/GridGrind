package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.DifferentialBorderPatch;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for protocol-facing conditional-formatting authoring inputs and conversion. */
class ConditionalFormattingInputTest {
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
}
