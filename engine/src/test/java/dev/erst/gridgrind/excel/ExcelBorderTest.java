package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellBorderPatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for CellBorderPatch and BorderSidePatch record construction. */
class ExcelBorderTest {
  @Test
  void constructsBorderPatchesWithDefaultsAndOverrides() {
    CellBorderPatch border =
        new CellBorderPatch(
            Optional.of(new BorderSidePatch(ExcelBorderStyle.THIN)),
            Optional.empty(),
            Optional.of(new BorderSidePatch(ExcelBorderStyle.DOUBLE)),
            Optional.empty(),
            Optional.empty());

    assertEquals(Optional.of(ExcelBorderStyle.THIN), border.all().orElseThrow().style());
    assertEquals(Optional.of(ExcelBorderStyle.DOUBLE), border.right().orElseThrow().style());
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN),
        new CellBorderPatch(
                Optional.empty(),
                Optional.of(new BorderSidePatch(ExcelBorderStyle.THIN)),
                Optional.empty(),
                Optional.empty(),
                Optional.empty())
            .top()
            .orElseThrow()
            .style());
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN),
        new CellBorderPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.of(new BorderSidePatch(ExcelBorderStyle.THIN)),
                Optional.empty(),
                Optional.empty())
            .right()
            .orElseThrow()
            .style());
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN),
        new CellBorderPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(new BorderSidePatch(ExcelBorderStyle.THIN)),
                Optional.empty())
            .bottom()
            .orElseThrow()
            .style());
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN),
        new CellBorderPatch(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(new BorderSidePatch(ExcelBorderStyle.THIN)))
            .left()
            .orElseThrow()
            .style());
  }

  @Test
  void validatesBorderRequirements() {
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
  }

  @Test
  void validatesBorderSideColorRules() {
    BorderSidePatch colorOnly =
        new BorderSidePatch(Optional.empty(), Optional.of(ColorSpec.rgb(new RgbHex("#a1b2c3"))));

    assertEquals(Optional.empty(), colorOnly.style());
    assertEquals(Optional.of(ColorSpec.rgb(new RgbHex("#A1B2C3"))), colorOnly.color());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new BorderSidePatch(
                Optional.of(ExcelBorderStyle.NONE),
                Optional.of(ColorSpec.rgb(new RgbHex("#112233")))));
  }
}
