package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for ExcelNamedRangeTarget record construction and formula rendering. */
class ExcelNamedRangeTargetTest {
  @Test
  void canonicalizesReversedRanges() {
    ExcelNamedRangeTarget target = new ExcelNamedRangeTarget("Budget", "B4:A1");

    assertEquals("A1:B4", target.range());
    assertEquals("Budget!$A$1:Budget!$B$4", target.refersToFormula());
  }

  @Test
  void rendersAbsoluteNamedRangeFormulas() {
    assertEquals("B4", new ExcelNamedRangeTarget("Budget", "B4").range());
    assertEquals("Budget!$B$4", new ExcelNamedRangeTarget("Budget", "B4").refersToFormula());
    assertEquals("B4:C4", new ExcelNamedRangeTarget("Budget", "B4:C4").range());
    assertEquals(
        "Budget!$B$4:Budget!$C$4", new ExcelNamedRangeTarget("Budget", "B4:C4").refersToFormula());
    assertEquals(
        "Budget!$B$4:Budget!$C$5", new ExcelNamedRangeTarget("Budget", "B4:C5").refersToFormula());
  }

  @Test
  void validatesNamedRangeTargetInputs() {
    assertThrows(NullPointerException.class, () -> new ExcelNamedRangeTarget(null, "A1"));
    assertThrows(IllegalArgumentException.class, () -> new ExcelNamedRangeTarget(" ", "A1"));
    assertThrows(NullPointerException.class, () -> new ExcelNamedRangeTarget("Budget", null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelNamedRangeTarget("Budget", " "));
    assertThrows(
        InvalidRangeAddressException.class, () -> new ExcelNamedRangeTarget("Budget", "A1:"));
  }
}
