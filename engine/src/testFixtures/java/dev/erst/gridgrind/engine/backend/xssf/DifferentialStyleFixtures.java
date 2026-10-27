package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.model.style.BorderSidePatch;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.style.DifferentialBorderPatch;
import dev.erst.gridgrind.core.model.style.DifferentialBorderReport;
import dev.erst.gridgrind.core.model.style.DifferentialBorderSideReport;

/** Expected factual colors and borders for tests that author differential styles through POI. */
public final class DifferentialStyleFixtures {
  private DifferentialStyleFixtures() {}

  /** Selects the expected factual color from the authored fixture. */
  public static CellColorReport color(ColorSpec color) {
    return color == null
        ? null
        : switch (color) {
          case ColorSpec.Rgb value -> new CellColorReport.Rgb(value.rgb(), value.tint());
          case ColorSpec.Theme value -> new CellColorReport.Theme(value.theme(), value.tint());
          case ColorSpec.Indexed value ->
              new CellColorReport.Indexed(value.indexed(), value.tint());
        };
  }

  /** Selects the expected factual border from an authored fixture. */
  public static DifferentialBorderReport border(DifferentialBorderPatch border) {
    return border == null
        ? null
        : new DifferentialBorderReport(
            border.all().map(DifferentialStyleFixtures::side),
            border.top().map(DifferentialStyleFixtures::side),
            border.right().map(DifferentialStyleFixtures::side),
            border.bottom().map(DifferentialStyleFixtures::side),
            border.left().map(DifferentialStyleFixtures::side));
  }

  /** Selects the expected factual side from an authored fixture. */
  public static DifferentialBorderSideReport side(BorderSidePatch side) {
    return new DifferentialBorderSideReport(
        side.style(), side.color().map(DifferentialStyleFixtures::color));
  }
}
