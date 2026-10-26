package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellBorderPatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for CellBorderPatch and shared BorderSidePatch conversion and validation. */
class CellBorderPatchTest {
  @Test
  void convertsProtocolBorderPatchesIntoEngineBorders() {
    CellBorderPatch border =
        new CellBorderPatch(
            Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.THIN)),
            Optional.empty(),
            Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.DOUBLE)),
            Optional.empty(),
            Optional.empty());
    CellBorderPatch bottomAndLeftBorder =
        new CellBorderPatch(
            Optional.empty(),
            Optional.empty(),
            Optional.empty(),
            Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.DASHED)),
            Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.DOTTED)));
    var engineBorder = WorkbookCommandCellValueSpecConverter.toExcelBorder(border).orElseThrow();
    var bottomAndLeftEngineBorder =
        WorkbookCommandCellValueSpecConverter.toExcelBorder(bottomAndLeftBorder).orElseThrow();

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
    assertThrows(IllegalArgumentException.class, () -> new BorderSidePatch(null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellBorderPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
    BorderSidePatch colorOnly = new BorderSidePatch(null, ColorSpec.rgb(new RgbHex("#a1b2c3")));
    assertEquals(Optional.empty(), colorOnly.style());
    assertEquals(Optional.of(ColorSpec.rgb(new RgbHex("#A1B2C3"))), colorOnly.color());
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN),
        WorkbookCommandCellValueSpecConverter.toExcelBorder(
                new CellBorderPatch(
                    Optional.empty(),
                    Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.THIN)),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty()))
            .orElseThrow()
            .top()
            .orElseThrow()
            .style());
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN),
        WorkbookCommandCellValueSpecConverter.toExcelBorder(
                new CellBorderPatch(
                    Optional.empty(),
                    Optional.empty(),
                    Optional.ofNullable(new BorderSidePatch(ExcelBorderStyle.THIN)),
                    Optional.empty(),
                    Optional.empty()))
            .orElseThrow()
            .right()
            .orElseThrow()
            .style());
  }
}
