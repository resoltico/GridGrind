package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import org.junit.jupiter.api.Test;

/** Tests for CellBorderInput and CellBorderSideInput conversion and validation. */
class CellBorderInputTest {
  @Test
  void convertsProtocolBorderPatchesIntoEngineBorders() {
    CellBorderInput border =
        new CellBorderInput(
            new CellBorderSideInput(ExcelBorderStyle.THIN),
            null,
            new CellBorderSideInput(ExcelBorderStyle.DOUBLE),
            null,
            null);
    CellBorderInput bottomAndLeftBorder =
        new CellBorderInput(
            null,
            null,
            null,
            new CellBorderSideInput(ExcelBorderStyle.DASHED),
            new CellBorderSideInput(ExcelBorderStyle.DOTTED));

    assertEquals(ExcelBorderStyle.THIN, border.toExcelBorder().all().style());
    assertEquals(ExcelBorderStyle.DOUBLE, border.toExcelBorder().right().style());
    assertEquals(ExcelBorderStyle.DASHED, bottomAndLeftBorder.toExcelBorder().bottom().style());
    assertEquals(ExcelBorderStyle.DOTTED, bottomAndLeftBorder.toExcelBorder().left().style());
  }

  @Test
  void validatesBorderPatchRequirements() {
    assertThrows(NullPointerException.class, () -> new CellBorderSideInput(null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellBorderInput(null, null, null, null, null));
    assertEquals(
        ExcelBorderStyle.THIN,
        new CellBorderInput(null, new CellBorderSideInput(ExcelBorderStyle.THIN), null, null, null)
            .toExcelBorder()
            .top()
            .style());
    assertEquals(
        ExcelBorderStyle.THIN,
        new CellBorderInput(null, null, new CellBorderSideInput(ExcelBorderStyle.THIN), null, null)
            .toExcelBorder()
            .right()
            .style());
  }
}
