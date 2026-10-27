package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.engine.poi.CellBounds;
import org.junit.jupiter.api.Test;

/** Tests for CellBounds parsing of A1-notation cell and range addresses. */
class CellBoundsTest {
  @Test
  void parsesSingleAndRectangularRanges() {
    CellBounds single = CellBounds.parse("B3");
    CellBounds rectangle = CellBounds.parse("D5:B3");

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
    assertThrows(IllegalArgumentException.class, () -> new CellBounds(-1, 0, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new CellBounds(5, 3, 0, 0));
    assertThrows(IllegalArgumentException.class, () -> new CellBounds(0, 0, -1, 0));
    assertThrows(IllegalArgumentException.class, () -> new CellBounds(0, 0, 5, 3));
  }

  @Test
  void rejectsInvalidRanges() {
    assertThrows(NullPointerException.class, () -> CellBounds.parse(null));
    assertThrows(IllegalArgumentException.class, () -> CellBounds.parse(" "));

    InvalidRangeAddressException blankEndpoint =
        assertThrows(InvalidRangeAddressException.class, () -> CellBounds.parse("A1:"));
    assertEquals("A1:", blankEndpoint.range());

    InvalidRangeAddressException extraSeparator =
        assertThrows(InvalidRangeAddressException.class, () -> CellBounds.parse("A1:B2:C3"));
    assertEquals("A1:B2:C3", extraSeparator.range());

    InvalidRangeAddressException invalidCell =
        assertThrows(InvalidRangeAddressException.class, () -> CellBounds.parse("A1:?"));
    assertEquals("A1:?", invalidCell.range());
  }
}
