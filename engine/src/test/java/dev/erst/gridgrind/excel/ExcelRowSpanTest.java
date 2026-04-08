package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import org.junit.jupiter.api.Test;

/** Tests for ExcelRowSpan bounds and count semantics. */
class ExcelRowSpanTest {
  @Test
  void validatesBoundsAndCountsRows() {
    assertEquals(3, new ExcelRowSpan(2, 4).count());

    IllegalArgumentException firstNegative =
        assertThrows(IllegalArgumentException.class, () -> new ExcelRowSpan(-1, 0));
    assertTrue(firstNegative.getMessage().contains("firstRowIndex -1"));
    assertTrue(firstNegative.getMessage().contains("Excel row 1"));

    IllegalArgumentException lastNegative =
        assertThrows(IllegalArgumentException.class, () -> new ExcelRowSpan(0, -1));
    assertTrue(lastNegative.getMessage().contains("lastRowIndex -1"));
    assertTrue(lastNegative.getMessage().contains("Excel row 1"));

    IllegalArgumentException descending =
        assertThrows(IllegalArgumentException.class, () -> new ExcelRowSpan(2, 1));
    assertTrue(descending.getMessage().contains("lastRowIndex 1 (Excel row 2)"));
    assertTrue(descending.getMessage().contains("firstRowIndex 2 (Excel row 3)"));

    IllegalArgumentException firstOverflow =
        assertThrows(
            IllegalArgumentException.class,
            () -> new ExcelRowSpan(ExcelRowSpan.MAX_ROW_INDEX + 1, ExcelRowSpan.MAX_ROW_INDEX + 1));
    assertTrue(firstOverflow.getMessage().contains("firstRowIndex 1048576 (Excel row 1048577)"));
    assertTrue(firstOverflow.getMessage().contains("1048575 (Excel row 1048576)"));

    IllegalArgumentException lastOverflow =
        assertThrows(
            IllegalArgumentException.class,
            () -> new ExcelRowSpan(0, ExcelRowSpan.MAX_ROW_INDEX + 1));
    assertTrue(lastOverflow.getMessage().contains("lastRowIndex 1048576 (Excel row 1048577)"));
    assertTrue(lastOverflow.getMessage().contains("1048575 (Excel row 1048576)"));
  }
}
