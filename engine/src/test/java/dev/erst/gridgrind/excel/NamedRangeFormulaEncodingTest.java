package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.io.IOException;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

/** Named-range writes normalize coordinate spelling only at the POI boundary. */
class NamedRangeFormulaEncodingTest {
  @ParameterizedTest
  @CsvSource({
    "B4:A1,Budget!$A$1:$B$4",
    "B4,Budget!$B$4",
    "B4:C4,Budget!$B$4:$C$4",
    "B4:C5,Budget!$B$4:$C$5"
  })
  void writesCanonicalAbsoluteRangesWithoutChangingAuthoredCoordinates(
      String authored, String expected) throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
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
}
