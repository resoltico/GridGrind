package dev.erst.gridgrind.core.model.value;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.formula.ResolvedRawFormula;
import org.junit.jupiter.api.Test;

/** Checks invalid scalar, temporal and materialized formula operands. */
class ResolvedCellWriteValidationTest {
  @Test
  void permissiveObservedScalarsCannotBecomeWriteOperands() {
    assertFalse(ResolvedCellWrite.class.isAssignableFrom(CellValue.Text.class));
    assertFalse(ResolvedCellWrite.class.isAssignableFrom(CellValue.NumberValue.class));
    assertFalse(ResolvedCellWrite.class.isAssignableFrom(CellValue.ErrorValue.class));
    assertEquals("", new CellValue.Text("").value());
    assertEquals("#CIRCULAR_REF!", new CellValue.ErrorValue("#CIRCULAR_REF!").value());
    assertTrue(Double.isNaN(new CellValue.NumberValue(Double.NaN).value()));
    assertThrows(IllegalArgumentException.class, () -> new CellValueSpec.NumberValue(Double.NaN));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellValueSpec.NumberValue(Double.POSITIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class, () -> new CellValueSpec.ErrorValue("#CIRCULAR_REF!"));
    assertThrows(IllegalArgumentException.class, () -> new ResolvedCellText(""));
    assertThrows(IllegalArgumentException.class, () -> new ResolvedCellText(" "));
    assertThrows(IllegalArgumentException.class, () -> new ResolvedCellText("invalid\u0000text"));
    var oversized =
        assertThrows(IllegalArgumentException.class, () -> new ResolvedCellText("x".repeat(32768)));
    assertTrue(oversized.getMessage().contains("Excel cell text limit"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ResolvedFormula("WEBSERVICE(\"https://example.invalid\")"));
    var invalidXml =
        assertThrows(
            IllegalArgumentException.class, () -> new ResolvedFormula("invalid\u0000formula"));
    var originalXml = assertInstanceOf(IllegalArgumentException.class, invalidXml.getCause());
    assertTrue(originalXml.getMessage().contains("XML 1.0-forbidden"));
    assertEquals(32767, new ResolvedCellText("x".repeat(32767)).value().length());
  }

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
