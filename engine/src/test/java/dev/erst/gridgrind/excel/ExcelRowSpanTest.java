package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Tests for ExcelRowSpan bounds and count semantics. */
class ExcelRowSpanTest {
  @Test
  void validatesBoundsAndCountsRows() {
    assertEquals(3, new ExcelRowSpan(2, 4).count());

    assertThrows(IllegalArgumentException.class, () -> new ExcelRowSpan(-1, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelRowSpan(0, -1));
    assertThrows(IllegalArgumentException.class, () -> new ExcelRowSpan(2, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelRowSpan(ExcelRowSpan.MAX_ROW_INDEX + 1, ExcelRowSpan.MAX_ROW_INDEX + 1));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelRowSpan(0, ExcelRowSpan.MAX_ROW_INDEX + 1));
  }
}
