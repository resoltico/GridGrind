package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for ExcelCellStyle record construction. */
class ExcelCellStyleTest {
  @Test
  void buildsStylePatchesAndFactories() {
    ExcelCellStyle numberFormat = ExcelCellStyle.numberFormat("#,##0.00");
    ExcelCellStyle emphasis = ExcelCellStyle.emphasis(true, false);
    ExcelCellStyle alignment =
        ExcelCellStyle.alignment(ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.TOP);

    assertEquals("#,##0.00", numberFormat.numberFormat());
    assertTrue(emphasis.bold());
    assertFalse(emphasis.italic());
    assertEquals(ExcelHorizontalAlignment.CENTER, alignment.horizontalAlignment());
    assertEquals(ExcelVerticalAlignment.TOP, alignment.verticalAlignment());
  }

  @Test
  void rejectsBlankOrEmptyStyles() {
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellStyle(" ", null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellStyle(null, null, null, null, null, null));
    assertThrows(IllegalArgumentException.class, () -> ExcelCellStyle.alignment(null, null));
  }
}
