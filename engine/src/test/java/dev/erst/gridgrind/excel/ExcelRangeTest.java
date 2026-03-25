package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for ExcelRange parsing of A1-notation cell and range addresses. */
class ExcelRangeTest {
  @Test
  void parsesSingleAndRectangularRanges() {
    ExcelRange single = ExcelRange.parse("B3");
    ExcelRange rectangle = ExcelRange.parse("D5:B3");

    assertEquals(2, single.firstRow());
    assertEquals(2, single.lastRow());
    assertEquals(1, single.firstColumn());
    assertEquals(1, single.lastColumn());
    assertEquals(1, single.rowCount());
    assertEquals(1, single.columnCount());

    assertEquals(2, rectangle.firstRow());
    assertEquals(4, rectangle.lastRow());
    assertEquals(1, rectangle.firstColumn());
    assertEquals(3, rectangle.lastColumn());
    assertEquals(3, rectangle.rowCount());
    assertEquals(3, rectangle.columnCount());
  }

  @Test
  void rejectsInvalidConstructorArguments() {
    assertThrows(IllegalArgumentException.class, () -> new ExcelRange(-1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelRange(5, 3, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelRange(0, 0, -1, 0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelRange(0, 0, 5, 3));
  }

  @Test
  void rejectsInvalidRanges() {
    assertThrows(NullPointerException.class, () -> ExcelRange.parse(null));
    assertThrows(IllegalArgumentException.class, () -> ExcelRange.parse(" "));

    InvalidRangeAddressException blankEndpoint =
        assertThrows(InvalidRangeAddressException.class, () -> ExcelRange.parse("A1:"));
    assertEquals("A1:", blankEndpoint.range());

    InvalidRangeAddressException extraSeparator =
        assertThrows(InvalidRangeAddressException.class, () -> ExcelRange.parse("A1:B2:C3"));
    assertEquals("A1:B2:C3", extraSeparator.range());

    InvalidRangeAddressException invalidCell =
        assertThrows(InvalidRangeAddressException.class, () -> ExcelRange.parse("A1:?"));
    assertEquals("A1:?", invalidCell.range());
  }
}
