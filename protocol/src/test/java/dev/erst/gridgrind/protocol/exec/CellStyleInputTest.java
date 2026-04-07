package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelFillPattern;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.protocol.dto.CellAlignmentInput;
import dev.erst.gridgrind.protocol.dto.CellBorderInput;
import dev.erst.gridgrind.protocol.dto.CellBorderSideInput;
import dev.erst.gridgrind.protocol.dto.CellFillInput;
import dev.erst.gridgrind.protocol.dto.CellFontInput;
import dev.erst.gridgrind.protocol.dto.CellProtectionInput;
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
            new CellAlignmentInput(
                true, ExcelHorizontalAlignment.RIGHT, ExcelVerticalAlignment.CENTER, 45, 3),
            new CellFontInput(
                true,
                false,
                "Aptos",
                new FontHeightInput.Points(new BigDecimal("11.5")),
                "#00aa55",
                true,
                false),
            new CellFillInput(ExcelFillPattern.THIN_HORIZONTAL_BANDS, "#FFF2CC", "#DDEBF7"),
            new CellBorderInput(
                null, new CellBorderSideInput(ExcelBorderStyle.THICK, "#112233"), null, null, null),
            new CellProtectionInput(false, true));

    var engineStyle = WorkbookCommandConverter.toExcelCellStyle(style);
    assertEquals("#,##0.00", engineStyle.numberFormat());
    assertTrue(engineStyle.font().bold());
    assertFalse(engineStyle.font().italic());
    assertTrue(engineStyle.alignment().wrapText());
    assertEquals(ExcelHorizontalAlignment.RIGHT, engineStyle.alignment().horizontalAlignment());
    assertEquals(ExcelVerticalAlignment.CENTER, engineStyle.alignment().verticalAlignment());
    assertEquals(45, engineStyle.alignment().textRotation());
    assertEquals(3, engineStyle.alignment().indentation());
    assertEquals("Aptos", engineStyle.font().fontName());
    assertEquals(230, engineStyle.font().fontHeight().twips());
    assertEquals(new BigDecimal("11.5"), engineStyle.font().fontHeight().points());
    assertEquals("#00AA55", engineStyle.font().fontColor());
    assertTrue(engineStyle.font().underline());
    assertFalse(engineStyle.font().strikeout());
    assertEquals(ExcelFillPattern.THIN_HORIZONTAL_BANDS, engineStyle.fill().pattern());
    assertEquals("#FFF2CC", engineStyle.fill().foregroundColor());
    assertEquals("#DDEBF7", engineStyle.fill().backgroundColor());
    assertEquals(ExcelBorderStyle.THICK, engineStyle.border().top().style());
    assertEquals("#112233", engineStyle.border().top().color());
    assertFalse(engineStyle.protection().locked());
    assertTrue(engineStyle.protection().hiddenFormula());
  }

  @Test
  void convertsStylesWithoutAlignmentSettings() {
    CellStyleInput style =
        new CellStyleInput(
            null,
            new CellAlignmentInput(true, null, null, null, null),
            new CellFontInput(null, true, null, null, null, null, null),
            null,
            null,
            null);

    var engineStyle = WorkbookCommandConverter.toExcelCellStyle(style);
    assertNull(engineStyle.numberFormat());
    assertNull(engineStyle.alignment().horizontalAlignment());
    assertNull(engineStyle.alignment().verticalAlignment());
    assertTrue(engineStyle.font().italic());
    assertTrue(engineStyle.alignment().wrapText());
  }

  @Test
  void rejectsBlankOrEmptyStyles() {
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellStyleInput(" ", null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellStyleInput(null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellAlignmentInput(null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, " ", null, null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new FontHeightInput.Twips(0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, null, null, "#12", null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, null, null, " ", null, null));
    assertThrows(IllegalArgumentException.class, () -> new CellFillInput(null, null, null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellFillInput(ExcelFillPattern.SOLID, " ", null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellAlignmentInput(null, null, null, -1, null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellAlignmentInput(null, null, null, 181, null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellAlignmentInput(null, null, null, null, -1));
    assertThrows(
        IllegalArgumentException.class, () -> new CellAlignmentInput(null, null, null, null, 251));
    assertThrows(IllegalArgumentException.class, () -> new CellFillInput(null, null, "#AABBCC"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillInput(ExcelFillPattern.NONE, "#AABBCC", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillInput(ExcelFillPattern.NONE, null, "#AABBCC"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFillInput(ExcelFillPattern.SOLID, "#AABBCC", "#112233"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellBorderSideInput(ExcelBorderStyle.NONE, "#AABBCC"));
    assertThrows(IllegalArgumentException.class, () -> new CellProtectionInput(null, null));
  }

  @Test
  void acceptsSingleAttributeStyles() {
    assertNotNull(
        new CellStyleInput(
            null, new CellAlignmentInput(true, null, null, null, null), null, null, null, null));
    assertNotNull(
        new CellStyleInput(
            null,
            new CellAlignmentInput(null, ExcelHorizontalAlignment.CENTER, null, null, null),
            null,
            null,
            null,
            null));
    assertNotNull(
        new CellStyleInput(
            null,
            null,
            new CellFontInput(null, null, "Aptos", null, null, null, null),
            null,
            null,
            null));
    assertNotNull(
        new CellStyleInput(
            null,
            null,
            new CellFontInput(null, null, null, new FontHeightInput.Twips(260), null, null, null),
            null,
            null,
            null));
    assertNotNull(
        new CellStyleInput(
            null,
            null,
            new CellFontInput(null, null, null, null, "#AABBCC", null, null),
            null,
            null,
            null));
    assertNotNull(
        new CellStyleInput(
            null,
            null,
            new CellFontInput(null, null, null, null, null, true, null),
            null,
            null,
            null));
    assertNotNull(
        new CellStyleInput(
            null,
            null,
            new CellFontInput(null, null, null, null, null, null, true),
            null,
            null,
            null));
    assertNotNull(
        new CellStyleInput(
            null,
            null,
            null,
            new CellFillInput(ExcelFillPattern.SOLID, "#AABBCC", null),
            null,
            null));
    assertNotNull(
        new CellStyleInput(
            null,
            new CellAlignmentInput(null, null, ExcelVerticalAlignment.TOP, null, null),
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
            new CellBorderInput(
                null, null, null, null, new CellBorderSideInput(ExcelBorderStyle.THIN)),
            null));
    CellAlignmentInput boundedAlignment =
        new CellAlignmentInput(null, null, ExcelVerticalAlignment.BOTTOM, 180, 250);
    assertEquals(180, boundedAlignment.textRotation());
    assertEquals(250, boundedAlignment.indentation());
    CellFillInput implicitSolidFill = new CellFillInput(null, "#AABBCC", null);
    assertEquals("#AABBCC", implicitSolidFill.foregroundColor());
    assertNotNull(new CellFillInput(ExcelFillPattern.NONE, null, null));
    assertNotNull(new CellFillInput(ExcelFillPattern.THIN_FORWARD_DIAGONAL, "#AABBCC", "#112233"));
    CellFillInput backgroundOnlyPattern =
        new CellFillInput(ExcelFillPattern.BRICKS, null, "#112233");
    assertEquals("#112233", backgroundOnlyPattern.backgroundColor());
    assertNotNull(new CellProtectionInput(false, null));
    CellProtectionInput hiddenOnlyProtection = new CellProtectionInput(null, true);
    assertTrue(hiddenOnlyProtection.hiddenFormula());
    CellStyleInput protectionOnlyStyle =
        new CellStyleInput(null, null, null, null, null, new CellProtectionInput(true, null));
    assertTrue(protectionOnlyStyle.protection().locked());
  }
}
