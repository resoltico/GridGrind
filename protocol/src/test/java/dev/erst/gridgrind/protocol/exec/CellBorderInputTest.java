package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.protocol.dto.CellBorderInput;
import dev.erst.gridgrind.protocol.dto.CellBorderSideInput;
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

    assertEquals(
        ExcelBorderStyle.THIN, WorkbookCommandConverter.toExcelBorder(border).all().style());
    assertEquals(
        ExcelBorderStyle.DOUBLE, WorkbookCommandConverter.toExcelBorder(border).right().style());
    assertEquals(
        ExcelBorderStyle.DASHED,
        WorkbookCommandConverter.toExcelBorder(bottomAndLeftBorder).bottom().style());
    assertEquals(
        ExcelBorderStyle.DOTTED,
        WorkbookCommandConverter.toExcelBorder(bottomAndLeftBorder).left().style());
  }

  @Test
  void validatesBorderPatchRequirements() {
    assertThrows(NullPointerException.class, () -> new CellBorderSideInput(null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellBorderInput(null, null, null, null, null));
    assertEquals(
        ExcelBorderStyle.THIN,
        WorkbookCommandConverter.toExcelBorder(
                new CellBorderInput(
                    null, new CellBorderSideInput(ExcelBorderStyle.THIN), null, null, null))
            .top()
            .style());
    assertEquals(
        ExcelBorderStyle.THIN,
        WorkbookCommandConverter.toExcelBorder(
                new CellBorderInput(
                    null, null, new CellBorderSideInput(ExcelBorderStyle.THIN), null, null))
            .right()
            .style());
  }
}
