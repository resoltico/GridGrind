package dev.erst.gridgrind.core.model.style;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.protection.CellProtectionPatch;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelFillPattern;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Checks canonical style patch validation and single-attribute updates. */
class CellStylePatchTest {
  @Test
  void rejectsBlankOrEmptyStyles() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellStylePatch(
                Optional.ofNullable(" "),
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
            new CellAlignmentPatch(
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
                Optional.ofNullable(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(IllegalArgumentException.class, () -> new FontHeightSpec.Twips(0));
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
        () ->
            new CellFontPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.ofNullable(ColorSpec.rgb(new RgbHex("#12"))),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellFontPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.ofNullable(ColorSpec.rgb(new RgbHex(" "))),
                Optional.empty(),
                Optional.empty()));
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
        new CellStylePatch(
            Optional.empty(),
            Optional.ofNullable(
                new CellAlignmentPatch(
                    Optional.of(true),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty()));
    assertNotNull(
        new CellStylePatch(
            Optional.empty(),
            Optional.ofNullable(
                new CellAlignmentPatch(
                    Optional.empty(),
                    Optional.of(ExcelHorizontalAlignment.CENTER),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty()));
    assertNotNull(
        new CellStylePatch(
            Optional.empty(),
            Optional.empty(),
            Optional.ofNullable(
                new CellFontPatch(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.ofNullable("Aptos"),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty()));
    assertNotNull(
        new CellStylePatch(
            Optional.empty(),
            Optional.empty(),
            Optional.ofNullable(
                new CellFontPatch(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.ofNullable(new FontHeightSpec.Twips(260)),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty()));
    assertNotNull(
        new CellStylePatch(
            Optional.empty(),
            Optional.empty(),
            Optional.ofNullable(
                new CellFontPatch(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.ofNullable(ColorSpec.rgb(new RgbHex("#AABBCC"))),
                    Optional.empty(),
                    Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty()));
    assertNotNull(
        new CellStylePatch(
            Optional.empty(),
            Optional.empty(),
            Optional.ofNullable(
                new CellFontPatch(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of(true),
                    Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty()));
    assertNotNull(
        new CellStylePatch(
            Optional.empty(),
            Optional.empty(),
            Optional.ofNullable(
                new CellFontPatch(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of(true))),
            Optional.empty(),
            Optional.empty(),
            Optional.empty()));
    assertNotNull(
        new CellStylePatch(
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.ofNullable(
                CellFillSpec.patternForeground(
                    ExcelFillPattern.SOLID, ColorSpec.rgb(new RgbHex("#AABBCC")))),
            Optional.empty(),
            Optional.empty()));
    assertNotNull(
        new CellStylePatch(
            Optional.empty(),
            Optional.ofNullable(
                new CellAlignmentPatch(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of(ExcelVerticalAlignment.TOP),
                    Optional.empty(),
                    Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty()));
    assertNotNull(
        new CellStylePatch(
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.ofNullable(
                new CellBorderPatch(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.THIN)))),
            Optional.empty()));
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
        new CellStylePatch(
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.ofNullable(new CellProtectionPatch(Optional.of(true), Optional.empty())));
    assertTrue(protectionOnlyStyle.protection().orElseThrow().locked().orElseThrow());
  }
}
