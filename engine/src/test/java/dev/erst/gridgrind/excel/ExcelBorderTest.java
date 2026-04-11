package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for ExcelBorder and ExcelBorderSide record construction. */
class ExcelBorderTest {
  @Test
  void constructsBorderPatchesWithDefaultsAndOverrides() {
    ExcelBorder border =
        new ExcelBorder(
            new ExcelBorderSide(ExcelBorderStyle.THIN),
            null,
            new ExcelBorderSide(ExcelBorderStyle.DOUBLE),
            null,
            null);

    assertEquals(ExcelBorderStyle.THIN, border.all().style());
    assertEquals(ExcelBorderStyle.DOUBLE, border.right().style());
    assertEquals(
        ExcelBorderStyle.THIN,
        new ExcelBorder(null, new ExcelBorderSide(ExcelBorderStyle.THIN), null, null, null)
            .top()
            .style());
    assertEquals(
        ExcelBorderStyle.THIN,
        new ExcelBorder(null, null, new ExcelBorderSide(ExcelBorderStyle.THIN), null, null)
            .right()
            .style());
    assertEquals(
        ExcelBorderStyle.THIN,
        new ExcelBorder(null, null, null, new ExcelBorderSide(ExcelBorderStyle.THIN), null)
            .bottom()
            .style());
    assertEquals(
        ExcelBorderStyle.THIN,
        new ExcelBorder(null, null, null, null, new ExcelBorderSide(ExcelBorderStyle.THIN))
            .left()
            .style());
  }

  @Test
  void validatesBorderRequirements() {
    assertThrows(IllegalArgumentException.class, () -> new ExcelBorderSide(null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelBorder(null, null, null, null, null));
  }

  @Test
  void validatesBorderSideColorRules() {
    ExcelBorderSide colorOnly = new ExcelBorderSide(null, new ExcelColor("#a1b2c3"));

    assertNull(colorOnly.style());
    assertEquals(new ExcelColor("#A1B2C3"), colorOnly.color());
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelBorderSide(ExcelBorderStyle.NONE, new ExcelColor("#112233")));
  }
}
