package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellBorderSideInput;
import dev.erst.gridgrind.contract.dto.ColorInput;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for CellBorderInput and CellBorderSideInput conversion and validation. */
class CellBorderInputTest {
  @Test
  void convertsProtocolBorderPatchesIntoEngineBorders() {
    CellBorderInput border =
        new CellBorderInput(
            Optional.ofNullable(new CellBorderSideInput(ExcelBorderStyle.THIN)),
            Optional.empty(),
            Optional.ofNullable(new CellBorderSideInput(ExcelBorderStyle.DOUBLE)),
            Optional.empty(),
            Optional.empty());
    CellBorderInput bottomAndLeftBorder =
        new CellBorderInput(
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.ofNullable(new CellBorderSideInput(ExcelBorderStyle.DASHED)),
            Optional.ofNullable(new CellBorderSideInput(ExcelBorderStyle.DOTTED)));
    var engineBorder = WorkbookCommandConverter.toExcelBorder(border).orElseThrow();
    var bottomAndLeftEngineBorder =
        WorkbookCommandConverter.toExcelBorder(bottomAndLeftBorder).orElseThrow();

    assertEquals(Optional.of(ExcelBorderStyle.THIN), engineBorder.all().orElseThrow().style());
    assertEquals(Optional.of(ExcelBorderStyle.DOUBLE), engineBorder.right().orElseThrow().style());
    assertEquals(
        Optional.of(ExcelBorderStyle.DASHED),
        bottomAndLeftEngineBorder.bottom().orElseThrow().style());
    assertEquals(
        Optional.of(ExcelBorderStyle.DOTTED),
        bottomAndLeftEngineBorder.left().orElseThrow().style());
  }

  @Test
  void validatesBorderPatchRequirements() {
    assertThrows(IllegalArgumentException.class, () -> new CellBorderSideInput(null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellBorderInput(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    CellBorderSideInput colorOnly = new CellBorderSideInput(null, ColorInput.rgb("#a1b2c3"));
    assertEquals(Optional.empty(), colorOnly.style());
    assertEquals(Optional.of(ColorInput.rgb("#A1B2C3")), colorOnly.color());
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN),
        WorkbookCommandConverter.toExcelBorder(
                new CellBorderInput(
                    Optional.empty(),
                    Optional.ofNullable(new CellBorderSideInput(ExcelBorderStyle.THIN)),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty()))
            .orElseThrow()
            .top()
            .orElseThrow()
            .style());
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN),
        WorkbookCommandConverter.toExcelBorder(
                new CellBorderInput(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.ofNullable(new CellBorderSideInput(ExcelBorderStyle.THIN)),
                    Optional.empty(),
                    Optional.empty()))
            .orElseThrow()
            .right()
            .orElseThrow()
            .style());
  }
}
