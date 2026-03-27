package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import org.junit.jupiter.api.Test;

/** Tests for NamedRangeTarget record construction and engine conversion. */
class NamedRangeTargetTest {
  @Test
  void canonicalizesReversedRanges() {
    NamedRangeTarget target = new NamedRangeTarget("Budget", "B4:A1");

    assertEquals("A1:B4", target.range());
    assertEquals(new ExcelNamedRangeTarget("Budget", "A1:B4"), target.toExcelNamedRangeTarget());
  }

  @Test
  void convertsNamedRangeTargetToEngineType() {
    NamedRangeTarget target = new NamedRangeTarget("Budget", "B4:C5");

    assertEquals(new ExcelNamedRangeTarget("Budget", "B4:C5"), target.toExcelNamedRangeTarget());
  }

  @Test
  void validatesNamedRangeTargetInputs() {
    assertThrows(NullPointerException.class, () -> new NamedRangeTarget(null, "A1"));
    assertThrows(IllegalArgumentException.class, () -> new NamedRangeTarget(" ", "A1"));
    assertThrows(NullPointerException.class, () -> new NamedRangeTarget("Budget", null));
    assertThrows(IllegalArgumentException.class, () -> new NamedRangeTarget("Budget", " "));
  }
}
