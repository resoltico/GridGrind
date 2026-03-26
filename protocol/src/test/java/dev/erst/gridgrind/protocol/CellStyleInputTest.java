package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import java.math.BigDecimal;
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
            ExcelHorizontalAlignment.RIGHT,
            ExcelVerticalAlignment.CENTER,
            "Aptos",
            new FontHeightInput.Points(new BigDecimal("11.5")),
            "#00aa55",
            true,
            false,
            "#FFF2CC",
            new CellBorderInput(
                null, new CellBorderSideInput(ExcelBorderStyle.THICK), null, null, null));

    assertEquals("#,##0.00", style.toExcelCellStyle().numberFormat());
    assertTrue(style.toExcelCellStyle().bold());
    assertFalse(style.toExcelCellStyle().italic());
    assertTrue(style.toExcelCellStyle().wrapText());
    assertEquals(ExcelHorizontalAlignment.RIGHT, style.toExcelCellStyle().horizontalAlignment());
    assertEquals(ExcelVerticalAlignment.CENTER, style.toExcelCellStyle().verticalAlignment());
    assertEquals("Aptos", style.toExcelCellStyle().fontName());
    assertEquals(230, style.toExcelCellStyle().fontHeight().twips());
    assertEquals(new BigDecimal("11.5"), style.toExcelCellStyle().fontHeight().points());
    assertEquals("#00AA55", style.toExcelCellStyle().fontColor());
    assertTrue(style.toExcelCellStyle().underline());
    assertFalse(style.toExcelCellStyle().strikeout());
    assertEquals("#FFF2CC", style.toExcelCellStyle().fillColor());
    assertEquals(ExcelBorderStyle.THICK, style.toExcelCellStyle().border().top().style());
  }

  @Test
  void convertsStylesWithoutAlignmentSettings() {
    CellStyleInput style =
        new CellStyleInput(
            null, null, true, true, null, null, null, null, null, null, null, null, null);

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
        () ->
            new CellStyleInput(
                " ", null, null, null, null, null, null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellStyleInput(
                null, null, null, null, null, null, null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellStyleInput(
                null,
                null,
                null,
                null,
                null,
                null,
                null,
                new FontHeightInput.Twips(0),
                null,
                null,
                null,
                null,
                null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellStyleInput(
                null, null, null, null, null, null, null, null, "#12", null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellStyleInput(
                null, null, null, null, null, null, " ", null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellStyleInput(
                null, null, null, null, null, null, null, null, " ", null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellStyleInput(
                null, null, null, null, null, null, null, null, null, null, null, " ", null));
  }

  @Test
  void acceptsSingleAttributeStyles() {
    assertNotNull(
        new CellStyleInput(
            null, null, null, true, null, null, null, null, null, null, null, null, null));
    assertNotNull(
        new CellStyleInput(
            null,
            null,
            null,
            null,
            ExcelHorizontalAlignment.CENTER,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null));
    assertNotNull(
        new CellStyleInput(
            null, null, null, null, null, null, "Aptos", null, null, null, null, null, null));
    assertNotNull(
        new CellStyleInput(
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            new FontHeightInput.Twips(260),
            null,
            null,
            null,
            null,
            null));
    assertNotNull(
        new CellStyleInput(
            null, null, null, null, null, null, null, null, "#AABBCC", null, null, null, null));
    assertNotNull(
        new CellStyleInput(
            null, null, null, null, null, null, null, null, null, true, null, null, null));
    assertNotNull(
        new CellStyleInput(
            null, null, null, null, null, null, null, null, null, null, true, null, null));
    assertNotNull(
        new CellStyleInput(
            null, null, null, null, null, null, null, null, null, null, null, "#AABBCC", null));
    assertNotNull(
        new CellStyleInput(
            null,
            null,
            null,
            null,
            null,
            ExcelVerticalAlignment.TOP,
            null,
            null,
            null,
            null,
            null,
            null,
            null));
    assertNotNull(
        new CellStyleInput(
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            new CellBorderInput(
                null, null, null, null, new CellBorderSideInput(ExcelBorderStyle.THIN))));
  }
}
