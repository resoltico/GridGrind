package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.CellAlignmentInput;
import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellBorderSideInput;
import dev.erst.gridgrind.contract.dto.CellFillInput;
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
        styleInput(
            "#,##0.00",
            new CellAlignmentInput(
                Optional.of(true),
                Optional.of(ExcelHorizontalAlignment.RIGHT),
                Optional.of(ExcelVerticalAlignment.CENTER),
                Optional.of(45),
                Optional.of(3)),
            fontInput(
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
                Optional.empty(),
                Optional.ofNullable(
                    new CellBorderSideInput(ExcelBorderStyle.THICK, ColorInput.rgb("#112233"))),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()),
            new CellProtectionInput(Optional.of(false), Optional.of(true)));

    var engineStyle = WorkbookCommandConverter.toExcelCellStyle(style);
    assertEquals("#,##0.00", engineStyle.numberFormat().orElseThrow());
    assertTrue(engineStyle.font().orElseThrow().bold().orElseThrow());
    assertFalse(engineStyle.font().orElseThrow().italic().orElseThrow());
    assertTrue(engineStyle.alignment().orElseThrow().wrapText().orElseThrow());
    assertEquals(
        ExcelHorizontalAlignment.RIGHT,
        engineStyle.alignment().orElseThrow().horizontalAlignment().orElseThrow());
    assertEquals(
        ExcelVerticalAlignment.CENTER,
        engineStyle.alignment().orElseThrow().verticalAlignment().orElseThrow());
    assertEquals(45, engineStyle.alignment().orElseThrow().textRotation().orElseThrow());
    assertEquals(3, engineStyle.alignment().orElseThrow().indentation().orElseThrow());
    assertEquals("Aptos", engineStyle.font().orElseThrow().fontName().orElseThrow());
    assertEquals(230, engineStyle.font().orElseThrow().fontHeight().orElseThrow().twips());
    assertEquals(
        new BigDecimal("11.5"),
        engineStyle.font().orElseThrow().fontHeight().orElseThrow().points());
    assertEquals(
        ExcelColor.rgb("#00AA55"), engineStyle.font().orElseThrow().fontColor().orElseThrow());
    assertTrue(engineStyle.font().orElseThrow().underline().orElseThrow());
    assertFalse(engineStyle.font().orElseThrow().strikeout().orElseThrow());
    ExcelCellFill.PatternForegroundBackground fill =
        assertInstanceOf(
            ExcelCellFill.PatternForegroundBackground.class, engineStyle.fill().orElseThrow());
    assertEquals(ExcelFillPattern.THIN_HORIZONTAL_BANDS, fill.pattern());
    assertEquals(ExcelColor.rgb("#FFF2CC"), fill.foregroundColor());
    assertEquals(ExcelColor.rgb("#DDEBF7"), fill.backgroundColor());
    assertEquals(
        Optional.of(ExcelBorderStyle.THICK),
        engineStyle.border().orElseThrow().top().orElseThrow().style());
    assertEquals(
        Optional.of(ExcelColor.rgb("#112233")),
        engineStyle.border().orElseThrow().top().orElseThrow().color());
    assertFalse(engineStyle.protection().orElseThrow().locked().orElseThrow());
    assertTrue(engineStyle.protection().orElseThrow().hiddenFormula().orElseThrow());
  }

  @Test
  void convertsStylesWithoutAlignmentSettings() {
    CellStyleInput style =
        styleInput(
            null,
            new CellAlignmentInput(
                Optional.of(true),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()),
            fontInput(null, true, null, null, null, null, null),
            null,
            null,
            null);

    var engineStyle = WorkbookCommandConverter.toExcelCellStyle(style);
    assertEquals(Optional.empty(), engineStyle.numberFormat());
    assertEquals(Optional.empty(), engineStyle.alignment().orElseThrow().horizontalAlignment());
    assertEquals(Optional.empty(), engineStyle.alignment().orElseThrow().verticalAlignment());
    assertTrue(engineStyle.font().orElseThrow().italic().orElseThrow());
    assertTrue(engineStyle.alignment().orElseThrow().wrapText().orElseThrow());
  }

  @Test
  void rejectsBlankOrEmptyStyles() {
    assertThrows(
        IllegalArgumentException.class, () -> styleInput(" ", null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class, () -> styleInput(null, null, null, null, null, null));
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
        IllegalArgumentException.class, () -> fontInput(null, null, " ", null, null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new FontHeightInput.Twips(0));
    assertThrows(
        IllegalArgumentException.class, () -> fontInput(null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> fontInput(null, null, null, null, ColorInput.rgb("#12"), null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> fontInput(null, null, null, null, ColorInput.rgb(" "), null, null));
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
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellProtectionInput(Optional.empty(), Optional.empty()));
  }

  @Test
  void acceptsSingleAttributeStyles() {
    assertNotNull(
        styleInput(
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
        styleInput(
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
        styleInput(
            null, null, fontInput(null, null, "Aptos", null, null, null, null), null, null, null));
    assertNotNull(
        styleInput(
            null,
            null,
            fontInput(null, null, null, new FontHeightInput.Twips(260), null, null, null),
            null,
            null,
            null));
    assertNotNull(
        styleInput(
            null,
            null,
            fontInput(null, null, null, null, ColorInput.rgb("#AABBCC"), null, null),
            null,
            null,
            null));
    assertNotNull(
        styleInput(
            null, null, fontInput(null, null, null, null, null, true, null), null, null, null));
    assertNotNull(
        styleInput(
            null, null, fontInput(null, null, null, null, null, null, true), null, null, null));
    assertNotNull(
        styleInput(
            null,
            null,
            null,
            CellFillInput.patternForeground(ExcelFillPattern.SOLID, ColorInput.rgb("#AABBCC")),
            null,
            null));
    assertNotNull(
        styleInput(
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
        styleInput(
            null,
            null,
            null,
            null,
            new CellBorderInput(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.ofNullable(new CellBorderSideInput(ExcelBorderStyle.THIN))),
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
    assertNotNull(new CellProtectionInput(Optional.of(false), Optional.empty()));
    CellProtectionInput hiddenOnlyProtection =
        new CellProtectionInput(Optional.empty(), Optional.of(true));
    assertTrue(hiddenOnlyProtection.hiddenFormula().orElseThrow());
    CellStyleInput protectionOnlyStyle =
        styleInput(
            null,
            null,
            null,
            null,
            null,
            new CellProtectionInput(Optional.of(true), Optional.empty()));
    assertTrue(protectionOnlyStyle.protection().orElseThrow().locked().orElseThrow());
  }
}
