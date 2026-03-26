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
    assertThrows(NullPointerException.class, () -> new ExcelBorderSide(null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelBorder(null, null, null, null, null));
  }
}
