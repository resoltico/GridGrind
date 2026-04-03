package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.protocol.dto.BorderStyle;
import dev.erst.gridgrind.protocol.dto.CellBorderInput;
import dev.erst.gridgrind.protocol.dto.CellBorderSideInput;
import org.junit.jupiter.api.Test;

/** Tests for CellBorderInput and CellBorderSideInput conversion and validation. */
class CellBorderInputTest {
  @Test
  void convertsProtocolBorderPatchesIntoEngineBorders() {
    CellBorderInput border =
        new CellBorderInput(
            new CellBorderSideInput(BorderStyle.THIN),
            null,
            new CellBorderSideInput(BorderStyle.DOUBLE),
            null,
            null);
    CellBorderInput bottomAndLeftBorder =
        new CellBorderInput(
            null,
            null,
            null,
            new CellBorderSideInput(BorderStyle.DASHED),
            new CellBorderSideInput(BorderStyle.DOTTED));

    assertEquals(
        ExcelBorderStyle.THIN, DefaultGridGrindRequestExecutor.toExcelBorder(border).all().style());
    assertEquals(
        ExcelBorderStyle.DOUBLE,
        DefaultGridGrindRequestExecutor.toExcelBorder(border).right().style());
    assertEquals(
        ExcelBorderStyle.DASHED,
        DefaultGridGrindRequestExecutor.toExcelBorder(bottomAndLeftBorder).bottom().style());
    assertEquals(
        ExcelBorderStyle.DOTTED,
        DefaultGridGrindRequestExecutor.toExcelBorder(bottomAndLeftBorder).left().style());
  }

  @Test
  void validatesBorderPatchRequirements() {
    assertThrows(NullPointerException.class, () -> new CellBorderSideInput(null));
    assertThrows(
        IllegalArgumentException.class, () -> new CellBorderInput(null, null, null, null, null));
    assertEquals(
        ExcelBorderStyle.THIN,
        DefaultGridGrindRequestExecutor.toExcelBorder(
                new CellBorderInput(
                    null, new CellBorderSideInput(BorderStyle.THIN), null, null, null))
            .top()
            .style());
    assertEquals(
        ExcelBorderStyle.THIN,
        DefaultGridGrindRequestExecutor.toExcelBorder(
                new CellBorderInput(
                    null, null, new CellBorderSideInput(BorderStyle.THIN), null, null))
            .right()
            .style());
  }
}
