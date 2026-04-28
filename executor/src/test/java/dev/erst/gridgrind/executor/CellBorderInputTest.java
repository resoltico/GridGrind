package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellBorderSideInput;
import dev.erst.gridgrind.contract.dto.ColorInput;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
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
    var engineBorder = WorkbookCommandConverter.toExcelBorder(border).orElseThrow();
    var bottomAndLeftEngineBorder =
        WorkbookCommandConverter.toExcelBorder(bottomAndLeftBorder).orElseThrow();

    assertEquals(ExcelBorderStyle.THIN, engineBorder.all().style());
    assertEquals(ExcelBorderStyle.DOUBLE, engineBorder.right().style());
    assertEquals(ExcelBorderStyle.DASHED, bottomAndLeftEngineBorder.bottom().style());
    assertEquals(ExcelBorderStyle.DOTTED, bottomAndLeftEngineBorder.left().style());
  }

  @Test
  void validatesBorderPatchRequirements() {
    assertThrows(IllegalArgumentException.class, () -> new CellBorderSideInput(null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellBorderInput(null, null, null, null, null));
    CellBorderSideInput colorOnly = new CellBorderSideInput(null, ColorInput.rgb("#a1b2c3"));
    assertNull(colorOnly.style());
    assertEquals(ColorInput.rgb("#A1B2C3"), colorOnly.color());
    assertEquals(
        ExcelBorderStyle.THIN,
        WorkbookCommandConverter.toExcelBorder(
                new CellBorderInput(
                    null, new CellBorderSideInput(ExcelBorderStyle.THIN), null, null, null))
            .orElseThrow()
            .top()
            .style());
    assertEquals(
        ExcelBorderStyle.THIN,
        WorkbookCommandConverter.toExcelBorder(
                new CellBorderInput(
                    null, null, new CellBorderSideInput(ExcelBorderStyle.THIN), null, null))
            .orElseThrow()
            .right()
            .style());
  }
}
