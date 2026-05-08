package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for ExcelBorder and ExcelBorderSide record construction. */
class ExcelBorderTest {
  @Test
  void constructsBorderPatchesWithDefaultsAndOverrides() {
    ExcelBorder border =
        new ExcelBorder(
            Optional.of(new ExcelBorderSide(ExcelBorderStyle.THIN)),
            Optional.empty(),
            Optional.of(new ExcelBorderSide(ExcelBorderStyle.DOUBLE)),
            Optional.empty(),
            Optional.empty());

    assertEquals(Optional.of(ExcelBorderStyle.THIN), border.all().orElseThrow().style());
    assertEquals(Optional.of(ExcelBorderStyle.DOUBLE), border.right().orElseThrow().style());
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN),
        new ExcelBorder(
                Optional.empty(),
                Optional.of(new ExcelBorderSide(ExcelBorderStyle.THIN)),
                Optional.empty(),
                Optional.empty(),
                Optional.empty())
            .top()
            .orElseThrow()
            .style());
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN),
        new ExcelBorder(
                Optional.empty(),
                Optional.empty(),
                Optional.of(new ExcelBorderSide(ExcelBorderStyle.THIN)),
                Optional.empty(),
                Optional.empty())
            .right()
            .orElseThrow()
            .style());
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN),
        new ExcelBorder(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(new ExcelBorderSide(ExcelBorderStyle.THIN)),
                Optional.empty())
            .bottom()
            .orElseThrow()
            .style());
    assertEquals(
        Optional.of(ExcelBorderStyle.THIN),
        new ExcelBorder(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.of(new ExcelBorderSide(ExcelBorderStyle.THIN)))
            .left()
            .orElseThrow()
            .style());
  }

  @Test
  void validatesBorderRequirements() {
    assertThrows(IllegalArgumentException.class, () -> new ExcelBorderSide(null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelBorder(
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty(),
                Optional.empty()));
  }

  @Test
  void validatesBorderSideColorRules() {
    ExcelBorderSide colorOnly =
        new ExcelBorderSide(Optional.empty(), Optional.of(ExcelColor.rgb("#a1b2c3")));

    assertEquals(Optional.empty(), colorOnly.style());
    assertEquals(Optional.of(ExcelColor.rgb("#A1B2C3")), colorOnly.color());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelBorderSide(
                Optional.of(ExcelBorderStyle.NONE), Optional.of(ExcelColor.rgb("#112233"))));
  }
}
