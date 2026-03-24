package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import org.junit.jupiter.api.Test;

/** Tests for CellStyleInput record construction and ExcelCellStyle conversion. */
class CellStyleInputTest {
  @Test
  void validatesAndConvertsStylePatches() {
    CellStyleInput style =
        new CellStyleInput(
            "#,##0.00",
            true,
            false,
            true,
            CellStyleInput.HorizontalAlignmentInput.RIGHT,
            CellStyleInput.VerticalAlignmentInput.CENTER);

    assertEquals("#,##0.00", style.toExcelCellStyle().numberFormat());
    assertTrue(style.toExcelCellStyle().bold());
    assertFalse(style.toExcelCellStyle().italic());
    assertTrue(style.toExcelCellStyle().wrapText());
    assertEquals(ExcelHorizontalAlignment.RIGHT, style.toExcelCellStyle().horizontalAlignment());
    assertEquals(ExcelVerticalAlignment.CENTER, style.toExcelCellStyle().verticalAlignment());
  }

  @Test
  void convertsStylesWithoutAlignmentSettings() {
    CellStyleInput style = new CellStyleInput(null, null, true, true, null, null);

    assertNull(style.toExcelCellStyle().numberFormat());
    assertNull(style.toExcelCellStyle().horizontalAlignment());
    assertNull(style.toExcelCellStyle().verticalAlignment());
    assertTrue(style.toExcelCellStyle().italic());
    assertTrue(style.toExcelCellStyle().wrapText());
  }

  @Test
  void rejectsBlankOrEmptyStyles() {
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellStyleInput(" ", null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellStyleInput(null, null, null, null, null, null));
  }
}
