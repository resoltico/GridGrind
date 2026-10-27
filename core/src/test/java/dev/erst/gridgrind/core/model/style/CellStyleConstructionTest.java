package dev.erst.gridgrind.core.model.style;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.protection.CellProtectionPatch;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.math.BigDecimal;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for CellStylePatch record construction. */
class CellStyleConstructionTest {
  @Test
  void buildsNestedStylePatches() {
    CellStylePatch numberFormat =
        new CellStylePatch(
            Optional.of("#,##0.00"),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    CellStylePatch emphasis =
        new CellStylePatch(
            Optional.empty(),
            Optional.empty(),
            Optional.of(
                new CellFontPatch(
                    Optional.of(true),
                    Optional.of(false),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    CellStylePatch alignment =
        new CellStylePatch(
            Optional.empty(),
            Optional.of(
                new CellAlignmentPatch(
                    Optional.empty(),
                    Optional.ofNullable(ExcelHorizontalAlignment.CENTER),
                    Optional.ofNullable(ExcelVerticalAlignment.TOP),
                    Optional.empty(),
                    Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    CellStylePatch fontAndFill =
        new CellStylePatch(
            Optional.empty(),
            Optional.empty(),
            Optional.of(
                new CellFontPatch(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of("Aptos"),
                    Optional.of(new FontHeightSpec.Points(new BigDecimal("11.5"))),
                    Optional.of(ColorSpec.rgb(new RgbHex("#00AAFF"))),
                    Optional.of(true),
                    Optional.of(false))),
            Optional.of(
                CellFillSpec.patternForeground(
                    ExcelFillPattern.SOLID, ColorSpec.rgb(new RgbHex("#FFF2CC")))),
            Optional.of(
                new CellBorderPatch(
                    Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.THIN)),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())),
            Optional.empty());

    assertEquals("#,##0.00", numberFormat.numberFormat().orElseThrow());
    assertTrue(emphasis.font().orElseThrow().bold().orElseThrow());
    assertFalse(emphasis.font().orElseThrow().italic().orElseThrow());
    assertEquals(
        ExcelHorizontalAlignment.CENTER,
        alignment.alignment().orElseThrow().horizontalAlignment().orElseThrow());
    assertEquals(
        ExcelVerticalAlignment.TOP,
        alignment.alignment().orElseThrow().verticalAlignment().orElseThrow());
    assertEquals("Aptos", fontAndFill.font().orElseThrow().fontName().orElseThrow());
    assertEquals(230, fontAndFill.font().orElseThrow().fontHeight().orElseThrow().twips());
    assertEquals(
        new BigDecimal("11.5"),
        fontAndFill.font().orElseThrow().fontHeight().orElseThrow().points());
    assertEquals(
        ColorSpec.rgb(new RgbHex("#00AAFF")),
        fontAndFill.font().orElseThrow().fontColor().orElseThrow());
    assertTrue(fontAndFill.font().orElseThrow().underline().orElseThrow());
    assertFalse(fontAndFill.font().orElseThrow().strikeout().orElseThrow());
    assertEquals(
        ColorSpec.rgb(new RgbHex("#FFF2CC")),
        assertInstanceOf(CellFillSpec.PatternForeground.class, fontAndFill.fill().orElseThrow())
            .foregroundColor());
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN),
        fontAndFill.border().orElseThrow().all().orElseThrow().style());
  }

  @Test
  void rejectsBlankOrEmptyStyles() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellStylePatch(
                Optional.of(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellStylePatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFontPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.of(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(IllegalArgumentException.class, () -> FontHeightReport.fromTwips(0));
    assertThrows(
        IllegalArgumentException.class, () -> FontHeightReport.fromTwips(Short.MAX_VALUE + 1));
    assertThrows(IllegalArgumentException.class, () -> ColorSpec.rgb(new RgbHex("#12ab")));
    assertThrows(IllegalArgumentException.class, () -> ColorSpec.rgb(new RgbHex(" ")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellFillSpec.patternForeground(ExcelFillPattern.SOLID, ColorSpec.rgb(new RgbHex(" "))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellStylePatch(
                Optional.empty(),
                Optional.of(
                    new CellAlignmentPatch(
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty())),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
  }

  @Test
  void acceptsSingleAttributePatches() {
    CellStylePatch italicOnly =
        new CellStylePatch(
            Optional.empty(),
            Optional.empty(),
            Optional.of(
                new CellFontPatch(
                    Optional.empty(),
                    Optional.of(true),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    assertTrue(italicOnly.font().orElseThrow().italic().orElseThrow());

    CellStylePatch wrapTextOnly =
        new CellStylePatch(
            Optional.empty(),
            Optional.of(
                new CellAlignmentPatch(
                    Optional.of(true),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    assertTrue(wrapTextOnly.alignment().orElseThrow().wrapText().orElseThrow());

    CellStylePatch fontColorOnly =
        new CellStylePatch(
            Optional.empty(),
            Optional.empty(),
            Optional.of(
                new CellFontPatch(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of(ColorSpec.rgb(new RgbHex("#aa00cc"))),
                    Optional.empty(),
                    Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    assertEquals(
        ColorSpec.rgb(new RgbHex("#AA00CC")),
        fontColorOnly.font().orElseThrow().fontColor().orElseThrow());

    CellStylePatch fontHeightOnly =
        new CellStylePatch(
            Optional.empty(),
            Optional.empty(),
            Optional.of(
                new CellFontPatch(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of(new FontHeightSpec.Twips(230)),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    assertEquals(
        new BigDecimal("11.5"),
        fontHeightOnly.font().orElseThrow().fontHeight().orElseThrow().points());

    CellStylePatch fillColorOnly =
        new CellStylePatch(
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.of(
                CellFillSpec.patternForeground(
                    ExcelFillPattern.SOLID, ColorSpec.rgb(new RgbHex("#abc123")))),
            Optional.empty(),
            Optional.empty());
    assertEquals(
        ColorSpec.rgb(new RgbHex("#ABC123")),
        assertInstanceOf(CellFillSpec.PatternForeground.class, fillColorOnly.fill().orElseThrow())
            .foregroundColor());

    CellStylePatch borderOnly =
        new CellStylePatch(
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.of(
                new CellBorderPatch(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.THIN)))),
            Optional.empty());
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN),
        borderOnly.border().orElseThrow().left().orElseThrow().style());

    CellStylePatch verticalOnly =
        new CellStylePatch(
            Optional.empty(),
            Optional.of(
                new CellAlignmentPatch(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.ofNullable(ExcelVerticalAlignment.TOP),
                    Optional.empty(),
                    Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    assertTrue(verticalOnly.alignment().orElseThrow().horizontalAlignment().isEmpty());
    assertEquals(
        ExcelVerticalAlignment.TOP,
        verticalOnly.alignment().orElseThrow().verticalAlignment().orElseThrow());

    CellStylePatch protectionOnly =
        new CellStylePatch(
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.of(new CellProtectionPatch(Optional.of(true), Optional.empty())));
    assertTrue(protectionOnly.protection().orElseThrow().locked().orElseThrow());
    assertEquals(Optional.empty(), protectionOnly.protection().orElseThrow().hiddenFormula());
  }

  @Test
  void validatesAlignmentContracts() {
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

    CellAlignmentPatch alignment =
        new CellAlignmentPatch(
            Optional.empty(),
            Optional.empty(),
            Optional.of(ExcelVerticalAlignment.BOTTOM),
            Optional.of(180),
            Optional.of(250));
    assertEquals(180, alignment.textRotation().orElseThrow());
    assertEquals(250, alignment.indentation().orElseThrow());
    assertEquals(ExcelVerticalAlignment.BOTTOM, alignment.verticalAlignment().orElseThrow());
  }

  @Test
  void validatesAuthoredAndFactualFillContracts() {
    assertThrows(NullPointerException.class, () -> CellFillSpec.pattern(null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellFillSpec.patternForeground(
                ExcelFillPattern.NONE, ColorSpec.rgb(new RgbHex("#112233"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellFillSpec.patternBackground(
                ExcelFillPattern.NONE, ColorSpec.rgb(new RgbHex("#112233"))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            CellFillSpec.patternColors(
                ExcelFillPattern.SOLID,
                ColorSpec.rgb(new RgbHex("#112233")),
                ColorSpec.rgb(new RgbHex("#445566"))));
    assertThrows(
        NullPointerException.class,
        () -> CellFillSpec.patternBackground(null, ColorSpec.rgb(new RgbHex("#445566"))));
    assertThrows(NullPointerException.class, () -> CellFillReport.pattern(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> CellFillReport.patternForeground(ExcelFillPattern.NONE, rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () -> CellFillReport.patternBackground(ExcelFillPattern.NONE, rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () -> CellFillReport.patternColors(ExcelFillPattern.SOLID, rgb("#112233"), rgb("#445566")));

    CellFillSpec.PatternForegroundBackground patternedFill =
        CellFillSpec.patternColors(
            ExcelFillPattern.BRICKS,
            ColorSpec.rgb(new RgbHex("#aa00cc")),
            ColorSpec.rgb(new RgbHex("#00bb11")));
    CellFillReport.PatternOnly noFill = CellFillReport.pattern(ExcelFillPattern.NONE);
    CellFillReport.PatternForegroundBackground patternedFillReport =
        CellFillReport.patternColors(ExcelFillPattern.BRICKS, rgb("#aa00cc"), rgb("#00bb11"));
    assertEquals(ExcelFillPattern.NONE, noFill.pattern());
    assertEquals(ColorSpec.rgb(new RgbHex("#AA00CC")), patternedFill.foregroundColor());
    assertEquals(ColorSpec.rgb(new RgbHex("#00BB11")), patternedFill.backgroundColor());
    assertEquals(rgb("#AA00CC"), patternedFillReport.foregroundColor());
    assertEquals(rgb("#00BB11"), patternedFillReport.backgroundColor());
  }

  @Test
  void validatesFontAndProtectionContracts() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFontPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellProtectionPatch(Optional.empty(), Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellStylePatch(
                Optional.empty(),
                Optional.empty(),
                Optional.of(
                    new CellFontPatch(
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty())),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));

    CellFontPatch strikeoutOnly =
        new CellFontPatch(
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.of(true));
    CellProtectionPatch hiddenOnly = new CellProtectionPatch(Optional.empty(), Optional.of(false));
    assertTrue(strikeoutOnly.strikeout().orElseThrow());
    assertFalse(hiddenOnly.hiddenFormula().orElseThrow());
    assertEquals(Optional.empty(), hiddenOnly.locked());
  }

  private static CellColorReport rgb(String rgb) {
    return CellColorReport.rgb(new RgbHex(rgb));
  }
}
