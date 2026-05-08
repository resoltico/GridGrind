package dev.erst.gridgrind.excel;

import static dev.erst.gridgrind.excel.ExcelStyleTestAccess.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelFillPattern;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.util.Optional;
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
            Optional.empty(),
            Optional.empty(),
            Optional.of(
                new ExcelCellFont(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of("Aptos"),
                    Optional.of(ExcelFontHeight.fromPoints(new BigDecimal("11.5"))),
                    Optional.of(ExcelColor.rgb("#00AAFF")),
                    Optional.of(true),
                    Optional.of(false))),
            Optional.of(
                ExcelCellFill.patternForeground(ExcelFillPattern.SOLID, ExcelColor.rgb("#FFF2CC"))),
            Optional.of(
                new ExcelBorder(
                    Optional.ofNullable(new ExcelBorderSide(ExcelBorderStyle.THIN)),
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
        ExcelColor.rgb("#00AAFF"), fontAndFill.font().orElseThrow().fontColor().orElseThrow());
    assertTrue(fontAndFill.font().orElseThrow().underline().orElseThrow());
    assertFalse(fontAndFill.font().orElseThrow().strikeout().orElseThrow());
    assertEquals(ExcelColor.rgb("#FFF2CC"), fillForegroundColor(fontAndFill.fill().orElseThrow()));
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN),
        fontAndFill.border().orElseThrow().all().orElseThrow().style());
  }

  @Test
  void rejectsBlankOrEmptyStyles() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellStyle(
                Optional.of(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellStyle(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellFont(
                Optional.empty(),
                Optional.empty(),
                Optional.of(" "),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(IllegalArgumentException.class, () -> new ExcelFontHeight(0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelFontHeight(Short.MAX_VALUE + 1));
    assertThrows(IllegalArgumentException.class, () -> ExcelColor.rgb("#12ab"));
    assertThrows(IllegalArgumentException.class, () -> ExcelColor.rgb(" "));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelCellFill.patternForeground(ExcelFillPattern.SOLID, ExcelColor.rgb(" ")));
    assertThrows(IllegalArgumentException.class, () -> ExcelCellStyle.alignment(null, null));
  }

  @Test
  void acceptsSingleAttributePatches() {
    ExcelCellStyle italicOnly =
        new ExcelCellStyle(
            Optional.empty(),
            Optional.empty(),
            Optional.of(
                new ExcelCellFont(
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

    ExcelCellStyle wrapTextOnly =
        new ExcelCellStyle(
            Optional.empty(),
            Optional.of(
                new ExcelCellAlignment(
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

    ExcelCellStyle fontColorOnly =
        new ExcelCellStyle(
            Optional.empty(),
            Optional.empty(),
            Optional.of(
                new ExcelCellFont(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of(ExcelColor.rgb("#aa00cc")),
                    Optional.empty(),
                    Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    assertEquals(
        ExcelColor.rgb("#AA00CC"), fontColorOnly.font().orElseThrow().fontColor().orElseThrow());

    ExcelCellStyle fontHeightOnly =
        new ExcelCellStyle(
            Optional.empty(),
            Optional.empty(),
            Optional.of(
                new ExcelCellFont(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.of(new ExcelFontHeight(230)),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty())),
            Optional.empty(),
            Optional.empty(),
            Optional.empty());
    assertEquals(
        new BigDecimal("11.5"),
        fontHeightOnly.font().orElseThrow().fontHeight().orElseThrow().points());

    ExcelCellStyle fillColorOnly =
        new ExcelCellStyle(
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.of(
                ExcelCellFill.patternForeground(ExcelFillPattern.SOLID, ExcelColor.rgb("#abc123"))),
            Optional.empty(),
            Optional.empty());
    assertEquals(
        ExcelColor.rgb("#ABC123"), fillForegroundColor(fillColorOnly.fill().orElseThrow()));

    ExcelCellStyle borderOnly =
        new ExcelCellStyle(
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.of(
                new ExcelBorder(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.ofNullable(new ExcelBorderSide(ExcelBorderStyle.THIN)))),
            Optional.empty());
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN),
        borderOnly.border().orElseThrow().left().orElseThrow().style());

    ExcelCellStyle verticalOnly = ExcelCellStyle.alignment(null, ExcelVerticalAlignment.TOP);
    assertTrue(verticalOnly.alignment().orElseThrow().horizontalAlignment().isEmpty());
    assertEquals(
        ExcelVerticalAlignment.TOP,
        verticalOnly.alignment().orElseThrow().verticalAlignment().orElseThrow());

    ExcelCellStyle protectionOnly =
        new ExcelCellStyle(
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.of(new ExcelCellProtection(Optional.of(true), Optional.empty())));
    assertTrue(protectionOnly.protection().orElseThrow().locked().orElseThrow());
    assertEquals(Optional.empty(), protectionOnly.protection().orElseThrow().hiddenFormula());
  }

  @Test
  void validatesAlignmentContracts() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellAlignment(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellAlignment(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(-1),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellAlignment(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(181),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellAlignment(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(-1)));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellAlignment(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(251)));

    ExcelCellAlignment alignment =
        new ExcelCellAlignment(
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
  void validatesFillContractsForPatchesAndSnapshots() {
    assertThrows(NullPointerException.class, () -> ExcelCellFill.pattern(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelCellFill.patternForeground(ExcelFillPattern.NONE, ExcelColor.rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelCellFill.patternBackground(ExcelFillPattern.NONE, ExcelColor.rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ExcelCellFill.patternColors(
                ExcelFillPattern.SOLID, ExcelColor.rgb("#112233"), ExcelColor.rgb("#445566")));
    assertThrows(
        NullPointerException.class,
        () -> ExcelCellFill.patternBackground(null, ExcelColor.rgb("#445566")));
    assertThrows(NullPointerException.class, () -> ExcelCellFillSnapshot.pattern(null));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelCellFillSnapshot.patternForeground(ExcelFillPattern.NONE, rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () -> ExcelCellFillSnapshot.patternBackground(ExcelFillPattern.NONE, rgb("#112233")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            ExcelCellFillSnapshot.patternColors(
                ExcelFillPattern.SOLID, rgb("#112233"), rgb("#445566")));

    ExcelCellFill patternedFill =
        ExcelCellFill.patternColors(
            ExcelFillPattern.BRICKS, ExcelColor.rgb("#aa00cc"), ExcelColor.rgb("#00bb11"));
    ExcelCellFillSnapshot noFillSnapshot = ExcelCellFillSnapshot.pattern(ExcelFillPattern.NONE);
    ExcelCellFillSnapshot patternedSnapshot =
        ExcelCellFillSnapshot.patternColors(
            ExcelFillPattern.BRICKS, rgb("#aa00cc"), rgb("#00bb11"));
    assertEquals(ExcelFillPattern.NONE, fillPattern(noFillSnapshot));
    assertNull(fillForegroundColor(noFillSnapshot));
    assertNull(fillBackgroundColor(noFillSnapshot));
    assertEquals(ExcelColor.rgb("#AA00CC"), fillForegroundColor(patternedFill));
    assertEquals(ExcelColor.rgb("#00BB11"), fillBackgroundColor(patternedFill));
    assertEquals(rgb("#AA00CC"), fillForegroundColor(patternedSnapshot));
    assertEquals(rgb("#00BB11"), fillBackgroundColor(patternedSnapshot));
  }

  @Test
  void validatesFontAndProtectionContracts() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelCellFont(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellProtection(Optional.empty(), Optional.empty()));
    assertThrows(IllegalArgumentException.class, () -> ExcelCellStyle.emphasis(null, null));

    ExcelCellFont strikeoutOnly =
        new ExcelCellFont(
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.of(true));
    ExcelCellProtection hiddenOnly = new ExcelCellProtection(Optional.empty(), Optional.of(false));
    assertTrue(strikeoutOnly.strikeout().orElseThrow());
    assertFalse(hiddenOnly.hiddenFormula().orElseThrow());
    assertEquals(Optional.empty(), hiddenOnly.locked());
  }

  private static ExcelColorSnapshot rgb(String rgb) {
    return ExcelColorSnapshot.rgb(rgb);
  }
}
