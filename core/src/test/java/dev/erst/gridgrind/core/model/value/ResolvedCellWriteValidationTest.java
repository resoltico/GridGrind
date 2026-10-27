package dev.erst.gridgrind.core.model.value;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedRawFormula;
import org.junit.jupiter.api.Test;

/** Checks invalid scalar, temporal and materialized formula operands. */
class ResolvedCellWriteValidationTest {
  @Test
  void validatesNonNullAndNonBlankValueRequirements() {
    assertThrows(NullPointerException.class, () -> new CellValue.Text(null));
    assertThrows(IllegalArgumentException.class, () -> new CellValue.ErrorValue(null));
    assertThrows(IllegalArgumentException.class, () -> new CellValue.ErrorValue(" "));
    assertThrows(IllegalArgumentException.class, () -> new CellValue.ErrorValue("~CIRCULAR~REF~"));
    assertThrows(IllegalArgumentException.class, () -> new CellValueSpec.Date(null));
    assertThrows(IllegalArgumentException.class, () -> new CellValueSpec.DateTime(null));
    assertThrows(NullPointerException.class, () -> new ResolvedFormula(null));
    assertThrows(IllegalArgumentException.class, () -> new ResolvedFormula(" "));
    assertThrows(NullPointerException.class, () -> new ResolvedRawFormula(null));
    assertThrows(IllegalArgumentException.class, () -> new ResolvedRawFormula(" "));
  }
}
