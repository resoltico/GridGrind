package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.CellAlignmentInput;
import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellBorderSideInput;
import dev.erst.gridgrind.contract.dto.CellFillInput;
import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.CellProtectionInput;
import dev.erst.gridgrind.contract.dto.CellStyleInput;
import dev.erst.gridgrind.contract.dto.ColorInput;
import dev.erst.gridgrind.contract.dto.FontHeightInput;
import dev.erst.gridgrind.excel.ExcelCellFill;
import dev.erst.gridgrind.excel.ExcelColor;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for CellStyleInput record construction and ExcelCellStyle conversion. */
class CellStyleInputTest {
  @Test
  void validatesAndConvertsStylePatches() {
    CellStyleInput style =
        new CellStyleInput(
            "#,##0.00",
            new CellAlignmentInput(
                Optional.of(true),
                Optional.of(ExcelHorizontalAlignment.RIGHT),
                Optional.of(ExcelVerticalAlignment.CENTER),
                Optional.of(45),
                Optional.of(3)),
            new CellFontInput(
                true,
                false,
                "Aptos",
                new FontHeightInput.Points(new BigDecimal("11.5")),
                ColorInput.rgb("#00aa55"),
                true,
                false),
            CellFillInput.patternColors(
                ExcelFillPattern.THIN_HORIZONTAL_BANDS,
                ColorInput.rgb("#FFF2CC"),
                ColorInput.rgb("#DDEBF7")),
            new CellBorderInput(
                null,
                new CellBorderSideInput(ExcelBorderStyle.THICK, ColorInput.rgb("#112233")),
                null,
                null,
                null),
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
    assertEquals(ExcelColor.rgb("#00AA55"), engineStyle.font().fontColor());
    assertTrue(engineStyle.font().underline());
    assertFalse(engineStyle.font().strikeout());
    ExcelCellFill.PatternForegroundBackground fill =
        assertInstanceOf(ExcelCellFill.PatternForegroundBackground.class, engineStyle.fill());
    assertEquals(ExcelFillPattern.THIN_HORIZONTAL_BANDS, fill.pattern());
    assertEquals(ExcelColor.rgb("#FFF2CC"), fill.foregroundColor());
    assertEquals(ExcelColor.rgb("#DDEBF7"), fill.backgroundColor());
    assertEquals(ExcelBorderStyle.THICK, engineStyle.border().top().style());
    assertEquals(ExcelColor.rgb("#112233"), engineStyle.border().top().color());
    assertFalse(engineStyle.protection().locked());
    assertTrue(engineStyle.protection().hiddenFormula());
  }

  @Test
  void convertsStylesWithoutAlignmentSettings() {
    CellStyleInput style =
        new CellStyleInput(
            null,
            new CellAlignmentInput(
                Optional.of(true),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()),
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
        IllegalArgumentException.class,
        () ->
            new CellAlignmentInput(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, " ", null, null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new FontHeightInput.Twips(0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, null, null, ColorInput.rgb("#12"), null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellFontInput(null, null, null, null, ColorInput.rgb(" "), null, null));
    assertThrows(NullPointerException.class, () -> CellFillInput.pattern(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> CellFillInput.patternForeground(ExcelFillPattern.SOLID, ColorInput.rgb(" ")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellAlignmentInput(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(-1),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellAlignmentInput(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(181),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellAlignmentInput(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(-1)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellAlignmentInput(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(251)));
    assertThrows(
        NullPointerException.class,
        () -> CellFillInput.patternBackground(null, ColorInput.rgb("#AABBCC")));
    assertThrows(
        IllegalArgumentException.class,
        () -> CellFillInput.patternForeground(ExcelFillPattern.NONE, ColorInput.rgb("#AABBCC")));
    assertThrows(
        IllegalArgumentException.class,
        () -> CellFillInput.patternBackground(ExcelFillPattern.NONE, ColorInput.rgb("#AABBCC")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellFillInput.patternColors(
                ExcelFillPattern.SOLID, ColorInput.rgb("#AABBCC"), ColorInput.rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellBorderSideInput(ExcelBorderStyle.NONE, ColorInput.rgb("#AABBCC")));
    assertThrows(IllegalArgumentException.class, () -> new CellProtectionInput(null, null));
  }

  @Test
  void acceptsSingleAttributeStyles() {
    assertNotNull(
        new CellStyleInput(
            null,
            new CellAlignmentInput(
                Optional.of(true),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()),
            null,
            null,
            null,
            null));
    assertNotNull(
        new CellStyleInput(
            null,
            new CellAlignmentInput(
                Optional.empty(),
                Optional.of(ExcelHorizontalAlignment.CENTER),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()),
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
            new CellFontInput(null, null, null, null, ColorInput.rgb("#AABBCC"), null, null),
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
            CellFillInput.patternForeground(ExcelFillPattern.SOLID, ColorInput.rgb("#AABBCC")),
            null,
            null));
    assertNotNull(
        new CellStyleInput(
            null,
            new CellAlignmentInput(
                Optional.empty(),
                Optional.empty(),
                Optional.of(ExcelVerticalAlignment.TOP),
                Optional.empty(),
                Optional.empty()),
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
        new CellAlignmentInput(
            Optional.empty(),
            Optional.empty(),
            Optional.of(ExcelVerticalAlignment.BOTTOM),
            Optional.of(180),
            Optional.of(250));
    assertEquals(180, boundedAlignment.textRotation().orElseThrow());
    assertEquals(250, boundedAlignment.indentation().orElseThrow());
    CellFillInput.PatternForeground solidFill =
        assertInstanceOf(
            CellFillInput.PatternForeground.class,
            CellFillInput.patternForeground(ExcelFillPattern.SOLID, ColorInput.rgb("#AABBCC")));
    assertEquals(ColorInput.rgb("#AABBCC"), solidFill.foregroundColor());
    assertNotNull(CellFillInput.pattern(ExcelFillPattern.NONE));
    assertNotNull(
        CellFillInput.patternColors(
            ExcelFillPattern.THIN_FORWARD_DIAGONAL,
            ColorInput.rgb("#AABBCC"),
            ColorInput.rgb("#112233")));
    CellFillInput.PatternBackground backgroundOnlyPattern =
        assertInstanceOf(
            CellFillInput.PatternBackground.class,
            CellFillInput.patternBackground(ExcelFillPattern.BRICKS, ColorInput.rgb("#112233")));
    assertEquals(ColorInput.rgb("#112233"), backgroundOnlyPattern.backgroundColor());
    assertNotNull(new CellProtectionInput(false, null));
    CellProtectionInput hiddenOnlyProtection = new CellProtectionInput(null, true);
    assertTrue(hiddenOnlyProtection.hiddenFormula());
    CellStyleInput protectionOnlyStyle =
        new CellStyleInput(null, null, null, null, null, new CellProtectionInput(true, null));
    assertTrue(protectionOnlyStyle.protection().locked());
  }
}
