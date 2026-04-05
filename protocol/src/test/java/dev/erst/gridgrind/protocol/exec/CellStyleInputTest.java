package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.protocol.dto.CellBorderInput;
import dev.erst.gridgrind.protocol.dto.CellBorderSideInput;
import dev.erst.gridgrind.protocol.dto.CellStyleInput;
import dev.erst.gridgrind.protocol.dto.FontHeightInput;
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

    var engineStyle = WorkbookCommandConverter.toExcelCellStyle(style);
    assertEquals("#,##0.00", engineStyle.numberFormat());
    assertTrue(engineStyle.bold());
    assertFalse(engineStyle.italic());
    assertTrue(engineStyle.wrapText());
    assertEquals(ExcelHorizontalAlignment.RIGHT, engineStyle.horizontalAlignment());
    assertEquals(ExcelVerticalAlignment.CENTER, engineStyle.verticalAlignment());
    assertEquals("Aptos", engineStyle.fontName());
    assertEquals(230, engineStyle.fontHeight().twips());
    assertEquals(new BigDecimal("11.5"), engineStyle.fontHeight().points());
    assertEquals("#00AA55", engineStyle.fontColor());
    assertTrue(engineStyle.underline());
    assertFalse(engineStyle.strikeout());
    assertEquals("#FFF2CC", engineStyle.fillColor());
    assertEquals(ExcelBorderStyle.THICK, engineStyle.border().top().style());
  }

  @Test
  void convertsStylesWithoutAlignmentSettings() {
    CellStyleInput style =
        new CellStyleInput(
            null, null, true, true, null, null, null, null, null, null, null, null, null);

    var engineStyle = WorkbookCommandConverter.toExcelCellStyle(style);
    assertNull(engineStyle.numberFormat());
    assertNull(engineStyle.horizontalAlignment());
    assertNull(engineStyle.verticalAlignment());
    assertTrue(engineStyle.italic());
    assertTrue(engineStyle.wrapText());
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
