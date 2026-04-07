package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Tests for ExcelColumnSpan bounds and count semantics. */
class ExcelColumnSpanTest {
  @Test
  void validatesBoundsAndCountsColumns() {
    assertEquals(3, new ExcelColumnSpan(2, 4).count());

    assertThrows(IllegalArgumentException.class, () -> new ExcelColumnSpan(-1, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelColumnSpan(0, -1));
    assertThrows(IllegalArgumentException.class, () -> new ExcelColumnSpan(2, 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelColumnSpan(
                ExcelColumnSpan.MAX_COLUMN_INDEX + 1, ExcelColumnSpan.MAX_COLUMN_INDEX + 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelColumnSpan(0, ExcelColumnSpan.MAX_COLUMN_INDEX + 1));
  }
}
