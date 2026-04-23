package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import org.junit.jupiter.api.Test;

/** Tests for ExcelColumnSpan bounds and count semantics. */
class ExcelColumnSpanTest {
  @Test
  void validatesBoundsAndCountsColumns() {
    assertEquals(3, new ExcelColumnSpan(2, 4).count());

    IllegalArgumentException firstNegative =
        assertThrows(IllegalArgumentException.class, () -> new ExcelColumnSpan(-1, 0));
    assertTrue(firstNegative.getMessage().contains("firstColumnIndex -1"));
    assertTrue(firstNegative.getMessage().contains("Excel column A"));

    IllegalArgumentException lastNegative =
        assertThrows(IllegalArgumentException.class, () -> new ExcelColumnSpan(0, -1));
    assertTrue(lastNegative.getMessage().contains("lastColumnIndex -1"));
    assertTrue(lastNegative.getMessage().contains("Excel column A"));

    IllegalArgumentException descending =
        assertThrows(IllegalArgumentException.class, () -> new ExcelColumnSpan(2, 1));
    assertTrue(descending.getMessage().contains("lastColumnIndex 1 (Excel column B)"));
    assertTrue(descending.getMessage().contains("firstColumnIndex 2 (Excel column C)"));

    IllegalArgumentException firstOverflow =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new ExcelColumnSpan(
                    ExcelColumnSpan.MAX_COLUMN_INDEX + 1, ExcelColumnSpan.MAX_COLUMN_INDEX + 1));
    assertTrue(firstOverflow.getMessage().contains("firstColumnIndex 16384 (Excel column XFE)"));
    assertTrue(firstOverflow.getMessage().contains("16383 (Excel column XFD)"));

    IllegalArgumentException lastOverflow =
        assertThrows(
            IllegalArgumentException.class,
            () -> new ExcelColumnSpan(0, ExcelColumnSpan.MAX_COLUMN_INDEX + 1));
    assertTrue(lastOverflow.getMessage().contains("lastColumnIndex 16384 (Excel column XFE)"));
    assertTrue(lastOverflow.getMessage().contains("16383 (Excel column XFD)"));
  }
}
