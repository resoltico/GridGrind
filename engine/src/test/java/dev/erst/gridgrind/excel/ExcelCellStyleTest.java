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
            new ExcelCellFont(
                null,
                null,
                "Aptos",
                ExcelFontHeight.fromPoints(new BigDecimal("11.5")),
                "#00AAFF",
                true,
                false),
            new ExcelCellFill(ExcelFillPattern.SOLID, "#FFF2CC", null),
            new ExcelBorder(new ExcelBorderSide(ExcelBorderStyle.THIN), null, null, null, null),
            null);

    assertEquals("#,##0.00", numberFormat.numberFormat());
    assertTrue(emphasis.font().bold());
    assertFalse(emphasis.font().italic());
    assertEquals(ExcelHorizontalAlignment.CENTER, alignment.alignment().horizontalAlignment());
    assertEquals(ExcelVerticalAlignment.TOP, alignment.alignment().verticalAlignment());
    assertEquals("Aptos", fontAndFill.font().fontName());
    assertEquals(230, fontAndFill.font().fontHeight().twips());
    assertEquals(new BigDecimal("11.5"), fontAndFill.font().fontHeight().points());
    assertEquals("#00AAFF", fontAndFill.font().fontColor());
    assertTrue(fontAndFill.font().underline());
    assertFalse(fontAndFill.font().strikeout());
    assertEquals("#FFF2CC", fontAndFill.fill().foregroundColor());
    assertEquals(ExcelBorderStyle.THIN, fontAndFill.border().all().style());
  }

  @Test
  void rejectsBlankOrEmptyStyles() {
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellStyle(" ", null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellStyle(null, null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellFont(null, null, " ", null, null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelFontHeight(0));
    assertThrows(IllegalArgumentException.class, () -> new ExcelFontHeight(Short.MAX_VALUE + 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellFont(null, null, null, null, "#12ab", null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellFont(null, null, null, null, " ", null, null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelCellFill(ExcelFillPattern.SOLID, " ", null));
    assertThrows(IllegalArgumentException.class, () -> ExcelCellStyle.alignment(null, null));
  }

  @Test
  void acceptsSingleAttributePatches() {
    ExcelCellStyle italicOnly =
        new ExcelCellStyle(
            null,
            null,
            new ExcelCellFont(null, true, null, null, null, null, null),
            null,
            null,
            null);
    assertTrue(italicOnly.font().italic());

    ExcelCellStyle wrapTextOnly =
        new ExcelCellStyle(
            null, new ExcelCellAlignment(true, null, null, null, null), null, null, null, null);
    assertTrue(wrapTextOnly.alignment().wrapText());

    ExcelCellStyle fontColorOnly =
        new ExcelCellStyle(
            null,
            null,
            new ExcelCellFont(null, null, null, null, "#aa00cc", null, null),
            null,
            null,
            null);
    assertEquals("#AA00CC", fontColorOnly.font().fontColor());

    ExcelCellStyle fontHeightOnly =
        new ExcelCellStyle(
            null,
            null,
            new ExcelCellFont(null, null, null, new ExcelFontHeight(230), null, null, null),
            null,
            null,
            null);
    assertEquals(new BigDecimal("11.5"), fontHeightOnly.font().fontHeight().points());

    ExcelCellStyle fillColorOnly =
        new ExcelCellStyle(
            null,
            null,
            null,
            new ExcelCellFill(ExcelFillPattern.SOLID, "#abc123", null),
            null,
            null);
    assertEquals("#ABC123", fillColorOnly.fill().foregroundColor());

    ExcelCellStyle borderOnly =
        new ExcelCellStyle(
            null,
            null,
            null,
            null,
            new ExcelBorder(null, null, null, null, new ExcelBorderSide(ExcelBorderStyle.THIN)),
            null);
    assertEquals(ExcelBorderStyle.THIN, borderOnly.border().left().style());

    ExcelCellStyle verticalOnly = ExcelCellStyle.alignment(null, ExcelVerticalAlignment.TOP);
    assertNull(verticalOnly.alignment().horizontalAlignment());
    assertEquals(ExcelVerticalAlignment.TOP, verticalOnly.alignment().verticalAlignment());

    ExcelCellStyle protectionOnly =
        new ExcelCellStyle(null, null, null, null, null, new ExcelCellProtection(true, null));
    assertTrue(protectionOnly.protection().locked());
    assertNull(protectionOnly.protection().hiddenFormula());
  }

  @Test
  void validatesAlignmentContracts() {
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelCellAlignment(null, null, null, null, null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelCellAlignment(null, null, null, -1, null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelCellAlignment(null, null, null, 181, null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelCellAlignment(null, null, null, null, -1));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelCellAlignment(null, null, null, null, 251));

    ExcelCellAlignment alignment =
        new ExcelCellAlignment(null, null, ExcelVerticalAlignment.BOTTOM, 180, 250);
    assertEquals(180, alignment.textRotation());
    assertEquals(250, alignment.indentation());
    assertEquals(ExcelVerticalAlignment.BOTTOM, alignment.verticalAlignment());
  }

  @Test
  void validatesFillContractsForPatchesAndSnapshots() {
    assertThrows(IllegalArgumentException.class, () -> new ExcelCellFill(null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellFill(ExcelFillPattern.NONE, "#112233", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellFill(ExcelFillPattern.NONE, null, "#112233"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellFill(ExcelFillPattern.SOLID, "#112233", "#445566"));
    assertThrows(IllegalArgumentException.class, () -> new ExcelCellFill(null, null, "#445566"));
    assertThrows(NullPointerException.class, () -> new ExcelCellFillSnapshot(null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellFillSnapshot(ExcelFillPattern.NONE, "#112233", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellFillSnapshot(ExcelFillPattern.NONE, null, "#112233"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellFillSnapshot(ExcelFillPattern.SOLID, "#112233", "#445566"));

    ExcelCellFill patternedFill = new ExcelCellFill(ExcelFillPattern.BRICKS, "#aa00cc", "#00bb11");
    ExcelCellFillSnapshot noFillSnapshot =
        new ExcelCellFillSnapshot(ExcelFillPattern.NONE, null, null);
    ExcelCellFillSnapshot patternedSnapshot =
        new ExcelCellFillSnapshot(ExcelFillPattern.BRICKS, "#aa00cc", "#00bb11");
    assertEquals(ExcelFillPattern.NONE, noFillSnapshot.pattern());
    assertNull(noFillSnapshot.foregroundColor());
    assertNull(noFillSnapshot.backgroundColor());
    assertEquals("#AA00CC", patternedFill.foregroundColor());
    assertEquals("#00BB11", patternedFill.backgroundColor());
    assertEquals("#AA00CC", patternedSnapshot.foregroundColor());
    assertEquals("#00BB11", patternedSnapshot.backgroundColor());
  }

  @Test
  void validatesFontAndProtectionContracts() {
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelCellFont(null, null, null, null, null, null, null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelCellProtection(null, null));
    assertThrows(IllegalArgumentException.class, () -> ExcelCellStyle.emphasis(null, null));

    ExcelCellFont strikeoutOnly = new ExcelCellFont(null, null, null, null, null, null, true);
    ExcelCellProtection hiddenOnly = new ExcelCellProtection(null, false);
    assertTrue(strikeoutOnly.strikeout());
    assertFalse(hiddenOnly.hiddenFormula());
    assertNull(hiddenOnly.locked());
  }
}
