package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.io.IOException;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/**
 * Named-range writes consume canonical geometry while formula targets retain exact authored text.
 */
class NamedRangeFormulaEncodingTest {
  @ParameterizedTest
  @CsvSource({
    "A1:B4,Budget!$A$1:$B$4",
    "B4,Budget!$B$4",
    "B4:C4,Budget!$B$4:$C$4",
    "B4:C5,Budget!$B$4:$C$5"
  })
  void writesAbsoluteFormulasFromCanonicalConcreteRanges(String authored, String expected)
      throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      var target = new NamedRangeTarget.Range(new SheetName("Budget"), new RangeRef(authored));
      workbook
          .names()
          .setNamedRange(
              new StructuredMutationAction.SetNamedRange(
                  new DefinedName("BudgetRange"), new NamedRangeScope.Workbook(), target));
      assertEquals(authored, target.range().value());
      assertEquals(expected, workbook.xssfWorkbook().getName("BudgetRange").getRefersToFormula());
    }
  }

  @Test
  void reversedConcreteRectanglesRejectWhileOpaqueFormulaOrderRemainsExact() throws IOException {
    assertThrows(
        dev.erst.gridgrind.core.problem.InvalidRequestException.class, () -> new RangeRef("B4:A1"));
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      String authored = "Budget!$B$4:$A$1";
      var formula = new NamedRangeTarget.Formula(authored);
      workbook
          .names()
          .setNamedRange(
              new StructuredMutationAction.SetNamedRange(
                  new DefinedName("FormulaRange"), new NamedRangeScope.Workbook(), formula));
      assertEquals(authored, formula.formula());
      assertEquals(authored, workbook.xssfWorkbook().getName("FormulaRange").getRefersToFormula());
    }
  }
}
