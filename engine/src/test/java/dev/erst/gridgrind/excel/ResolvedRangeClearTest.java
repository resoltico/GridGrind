package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Clear operations accept every range selector declared in their core contract. */
class ResolvedRangeClearTest {
  @Test
  void validationClearAcceptsRectangularWindow() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      var sheet = workbook.getOrCreateSheet("Rules");
      sheet
          .metadata()
          .setDataValidation(
              "A1:B3",
              new ResolvedDataValidation(
                  new DataValidationRuleSpec.ExplicitList(List.of("yes", "no")),
                  true,
                  false,
                  Optional.empty(),
                  Optional.empty()));
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new RangeSelector.RectangularWindow(
                  new SheetName("Rules"), new CellAddress("A1"), 3, 2),
              new StructuredMutationAction.ClearDataValidations()));
      assertEquals(0, sheet.metadata().dataValidations(new ExcelRangeSelection.All()).size());
    }
  }

  @Test
  void conditionalFormattingClearAcceptsRectangularWindow() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      var sheet = workbook.getOrCreateSheet("Rules");
      sheet
          .metadata()
          .setConditionalFormatting(
              List.of(new RangeRef("A1:B3")),
              new ConditionalFormattingSpec(
                  List.of(
                      new ConditionalFormattingRuleSpec.FormulaRule(
                          "A1>0", true, Optional.empty()))));
      ResolvedMutationExecutor.apply(
          workbook,
          new ResolvedStep(
              new RangeSelector.RectangularWindow(
                  new SheetName("Rules"), new CellAddress("A1"), 3, 2),
              new StructuredMutationAction.ClearConditionalFormatting()));
      assertEquals(0, sheet.metadata().conditionalFormatting(new ExcelRangeSelection.All()).size());
    }
  }
}
