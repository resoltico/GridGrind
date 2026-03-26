package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.math.BigDecimal;
import org.junit.jupiter.api.Test;

/** Tests for ExcelCellStyle record construction. */
class ExcelCellStyleTest {
  @Test
  void buildsStylePatchesAndFactories() {
    ExcelCellStyle numberFormat = ExcelCellStyle.numberFormat("#,##0.00");
    ExcelCellStyle emphasis = ExcelCellStyle.emphasis(true, false);
    ExcelCellStyle alignment =
        ExcelCellStyle.alignment(ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.TOP);
    ExcelCellStyle fontAndFill =
        new ExcelCellStyle(
            null,
            null,
            null,
            null,
            null,
            null,
            "Aptos",
            ExcelFontHeight.fromPoints(new BigDecimal("11.5")),
            "#00AAFF",
            true,
            false,
            "#FFF2CC",
            new ExcelBorder(new ExcelBorderSide(ExcelBorderStyle.THIN), null, null, null, null));

    assertEquals("#,##0.00", numberFormat.numberFormat());
    assertTrue(emphasis.bold());
    assertFalse(emphasis.italic());
    assertEquals(ExcelHorizontalAlignment.CENTER, alignment.horizontalAlignment());
    assertEquals(ExcelVerticalAlignment.TOP, alignment.verticalAlignment());
    assertEquals("Aptos", fontAndFill.fontName());
    assertEquals(230, fontAndFill.fontHeight().twips());
    assertEquals(new BigDecimal("11.5"), fontAndFill.fontHeight().points());
    assertEquals("#00AAFF", fontAndFill.fontColor());
    assertTrue(fontAndFill.underline());
    assertFalse(fontAndFill.strikeout());
    assertEquals("#FFF2CC", fontAndFill.fillColor());
    assertEquals(ExcelBorderStyle.THIN, fontAndFill.border().all().style());
  }

  @Test
  void rejectsBlankOrEmptyStyles() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellStyle(
                " ", null, null, null, null, null, null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellStyle(
                null, null, null, null, null, null, null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellStyle(
                null, null, null, null, null, null, " ", null, null, null, null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelFontHeight(0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelFontHeight(Short.MAX_VALUE + 1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellStyle(
                null, null, null, null, null, null, null, null, "#12ab", null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellStyle(
                null, null, null, null, null, null, null, null, " ", null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellStyle(
                null, null, null, null, null, null, null, null, null, null, null, " ", null));
    assertThrows(IllegalArgumentException.class, () -> ExcelCellStyle.alignment(null, null));
  }

  @Test
  void acceptsSingleAttributePatches() {
    ExcelCellStyle italicOnly =
        new ExcelCellStyle(
            null, null, true, null, null, null, null, null, null, null, null, null, null);
    assertTrue(italicOnly.italic());

    ExcelCellStyle wrapTextOnly =
        new ExcelCellStyle(
            null, null, null, true, null, null, null, null, null, null, null, null, null);
    assertTrue(wrapTextOnly.wrapText());

    ExcelCellStyle fontColorOnly =
        new ExcelCellStyle(
            null, null, null, null, null, null, null, null, "#aa00cc", null, null, null, null);
    assertEquals("#AA00CC", fontColorOnly.fontColor());

    ExcelCellStyle fontHeightOnly =
        new ExcelCellStyle(
            null,
            null,
            null,
            null,
            null,
            null,
            null,
            new ExcelFontHeight(230),
            null,
            null,
            null,
            null,
            null);
    assertEquals(new BigDecimal("11.5"), fontHeightOnly.fontHeight().points());

    ExcelCellStyle fillColorOnly =
        new ExcelCellStyle(
            null, null, null, null, null, null, null, null, null, null, null, "#abc123", null);
    assertEquals("#ABC123", fillColorOnly.fillColor());

    ExcelCellStyle borderOnly =
        new ExcelCellStyle(
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
            new ExcelBorder(null, null, null, null, new ExcelBorderSide(ExcelBorderStyle.THIN)));
    assertEquals(ExcelBorderStyle.THIN, borderOnly.border().left().style());

    ExcelCellStyle verticalOnly = ExcelCellStyle.alignment(null, ExcelVerticalAlignment.TOP);
    assertNull(verticalOnly.horizontalAlignment());
    assertEquals(ExcelVerticalAlignment.TOP, verticalOnly.verticalAlignment());
  }
}
