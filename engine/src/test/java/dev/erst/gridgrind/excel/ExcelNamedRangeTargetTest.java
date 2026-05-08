package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for ExcelNamedRangeTarget record construction and formula rendering. */
class ExcelNamedRangeTargetTest {
  @Test
  void canonicalizesReversedRanges() {
    ExcelNamedRangeTarget.Range target =
        (ExcelNamedRangeTarget.Range) ExcelNamedRangeTarget.range("Budget", "B4:A1");

    assertEquals("A1:B4", target.range());
    assertEquals("Budget!$A$1:$B$4", target.refersToFormula());
  }

  @Test
  void rendersAbsoluteNamedRangeFormulas() {
    assertEquals(
        "B4", ((ExcelNamedRangeTarget.Range) ExcelNamedRangeTarget.range("Budget", "B4")).range());
    assertEquals("Budget!$B$4", ExcelNamedRangeTarget.range("Budget", "B4").refersToFormula());
    assertEquals(
        "B4:C4",
        ((ExcelNamedRangeTarget.Range) ExcelNamedRangeTarget.range("Budget", "B4:C4")).range());
    assertEquals(
        "Budget!$B$4:$C$4", ExcelNamedRangeTarget.range("Budget", "B4:C4").refersToFormula());
    assertEquals(
        "Budget!$B$4:$C$5", ExcelNamedRangeTarget.range("Budget", "B4:C5").refersToFormula());
  }

  @Test
  void validatesNamedRangeTargetInputs() {
    assertThrows(NullPointerException.class, () -> ExcelNamedRangeTarget.range(null, "A1"));
    assertThrows(IllegalArgumentException.class, () -> ExcelNamedRangeTarget.range(" ", "A1"));
    assertThrows(NullPointerException.class, () -> ExcelNamedRangeTarget.range("Budget", null));
    assertThrows(IllegalArgumentException.class, () -> ExcelNamedRangeTarget.range("Budget", " "));
    assertThrows(IllegalArgumentException.class, () -> ExcelNamedRangeTarget.formula(" "));
    assertThrows(
        InvalidRangeAddressException.class, () -> ExcelNamedRangeTarget.range("Budget", "A1:"));
    assertDoesNotThrow(() -> ExcelNamedRangeTarget.range("A".repeat(31), "A1"));
    assertThrows(
        IllegalArgumentException.class, () -> ExcelNamedRangeTarget.range("A".repeat(32), "A1"));
  }
}
