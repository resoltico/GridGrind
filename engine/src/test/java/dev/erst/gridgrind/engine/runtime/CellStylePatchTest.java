package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.protection.CellProtectionPatch;
import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellAlignmentPatch;
import dev.erst.gridgrind.core.model.style.CellBorderPatch;
import dev.erst.gridgrind.core.model.style.CellFillSpec;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.FontHeightSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.excel.ExcelCellFill;
import dev.erst.gridgrind.excel.ExcelColor;
import java.math.BigDecimal;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for CellStylePatch record construction and ExcelCellStyle conversion. */
class CellStylePatchTest {
  @Test
  void validatesAndConvertsStylePatches() {
    CellStylePatch style =
        styleInput(
            "#,##0.00",
            new CellAlignmentPatch(
                Optional.of(true),
                Optional.of(ExcelHorizontalAlignment.RIGHT),
                Optional.of(ExcelVerticalAlignment.CENTER),
                Optional.of(45),
                Optional.of(3)),
            fontInput(
                true,
                false,
                "Aptos",
                new FontHeightSpec.Points(new BigDecimal("11.5")),
                ColorSpec.rgb(new RgbHex("#00aa55")),
                true,
                false),
            CellFillSpec.patternColors(
                ExcelFillPattern.THIN_HORIZONTAL_BANDS,
                ColorSpec.rgb(new RgbHex("#FFF2CC")),
                ColorSpec.rgb(new RgbHex("#DDEBF7"))),
            new CellBorderPatch(
                Optional.empty(),
                Optional.ofNullable(
                    new BorderSidePatch(
                        ExcelBorderStyle.THICK, ColorSpec.rgb(new RgbHex("#112233")))),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()),
            new CellProtectionPatch(Optional.of(false), Optional.of(true)));

    var engineStyle = WorkbookCommandCellValueSpecConverter.toExcelCellStyle(style);
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
    CellStylePatch style =
        styleInput(
            null,
            new CellAlignmentPatch(
                Optional.of(true),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()),
            fontInput(null, true, null, null, null, null, null),
            null,
            null,
            null);

    var engineStyle = WorkbookCommandCellValueSpecConverter.toExcelCellStyle(style);
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
            new CellAlignmentPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class, () -> fontInput(null, null, " ", null, null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new FontHeightSpec.Twips(0));
    assertThrows(
        IllegalArgumentException.class, () -> fontInput(null, null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> fontInput(null, null, null, null, ColorSpec.rgb(new RgbHex("#12")), null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> fontInput(null, null, null, null, ColorSpec.rgb(new RgbHex(" ")), null, null));
    assertThrows(NullPointerException.class, () -> CellFillSpec.pattern(null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellFillSpec.patternForeground(ExcelFillPattern.SOLID, ColorSpec.rgb(new RgbHex(" "))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellAlignmentPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(-1),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellAlignmentPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(181),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellAlignmentPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(-1)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellAlignmentPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(251)));
    assertThrows(
        NullPointerException.class,
        () -> CellFillSpec.patternBackground(null, ColorSpec.rgb(new RgbHex("#AABBCC"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellFillSpec.patternForeground(
                ExcelFillPattern.NONE, ColorSpec.rgb(new RgbHex("#AABBCC"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellFillSpec.patternBackground(
                ExcelFillPattern.NONE, ColorSpec.rgb(new RgbHex("#AABBCC"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellFillSpec.patternColors(
                ExcelFillPattern.SOLID,
                ColorSpec.rgb(new RgbHex("#AABBCC")),
                ColorSpec.rgb(new RgbHex("#112233"))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new BorderSidePatch(ExcelBorderStyle.NONE, ColorSpec.rgb(new RgbHex("#AABBCC"))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellProtectionPatch(Optional.empty(), Optional.empty()));
  }

  @Test
  void acceptsSingleAttributeStyles() {
    assertNotNull(
        styleInput(
            null,
            new CellAlignmentPatch(
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
            new CellAlignmentPatch(
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
            fontInput(null, null, null, new FontHeightSpec.Twips(260), null, null, null),
            null,
            null,
            null));
    assertNotNull(
        styleInput(
            null,
            null,
            fontInput(null, null, null, null, ColorSpec.rgb(new RgbHex("#AABBCC")), null, null),
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
            CellFillSpec.patternForeground(
                ExcelFillPattern.SOLID, ColorSpec.rgb(new RgbHex("#AABBCC"))),
            null,
            null));
    assertNotNull(
        styleInput(
            null,
            new CellAlignmentPatch(
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
            new CellBorderPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.THIN))),
            null));
    CellAlignmentPatch boundedAlignment =
        new CellAlignmentPatch(
            Optional.empty(),
            Optional.empty(),
            Optional.of(ExcelVerticalAlignment.BOTTOM),
            Optional.of(180),
            Optional.of(250));
    assertEquals(180, boundedAlignment.textRotation().orElseThrow());
    assertEquals(250, boundedAlignment.indentation().orElseThrow());
    CellFillSpec.PatternForeground solidFill =
        assertInstanceOf(
            CellFillSpec.PatternForeground.class,
            CellFillSpec.patternForeground(
                ExcelFillPattern.SOLID, ColorSpec.rgb(new RgbHex("#AABBCC"))));
    assertEquals(ColorSpec.rgb(new RgbHex("#AABBCC")), solidFill.foregroundColor());
    assertNotNull(CellFillSpec.pattern(ExcelFillPattern.NONE));
    assertNotNull(
        CellFillSpec.patternColors(
            ExcelFillPattern.THIN_FORWARD_DIAGONAL,
            ColorSpec.rgb(new RgbHex("#AABBCC")),
            ColorSpec.rgb(new RgbHex("#112233"))));
    CellFillSpec.PatternBackground backgroundOnlyPattern =
        assertInstanceOf(
            CellFillSpec.PatternBackground.class,
            CellFillSpec.patternBackground(
                ExcelFillPattern.BRICKS, ColorSpec.rgb(new RgbHex("#112233"))));
    assertEquals(ColorSpec.rgb(new RgbHex("#112233")), backgroundOnlyPattern.backgroundColor());
    assertNotNull(new CellProtectionPatch(Optional.of(false), Optional.empty()));
    CellProtectionPatch hiddenOnlyProtection =
        new CellProtectionPatch(Optional.empty(), Optional.of(true));
    assertTrue(hiddenOnlyProtection.hiddenFormula().orElseThrow());
    CellStylePatch protectionOnlyStyle =
        styleInput(
            null,
            null,
            null,
            null,
            null,
            new CellProtectionPatch(Optional.of(true), Optional.empty()));
    assertTrue(protectionOnlyStyle.protection().orElseThrow().locked().orElseThrow());
  }
}
