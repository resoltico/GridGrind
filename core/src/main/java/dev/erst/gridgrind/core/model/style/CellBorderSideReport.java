package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.ExcelBorderStyle;
import java.util.Objects;

/** Effective facts for one resolved cell-border side with explicit color semantics. */
@Union(discriminator = "type", catalog = true, displayOrder = 59)
public sealed interface CellBorderSideReport
    permits dev.erst.gridgrind.core.model.style.CellBorderSideReport.None,
        dev.erst.gridgrind.core.model.style.CellBorderSideReport.DefaultColor,
        dev.erst.gridgrind.core.model.style.CellBorderSideReport.Colored {
  /** Reports that this border side has no visible border. */
  @Variant(id = "NONE", summary = "Read one border side with no visible border.")
  record None() implements CellBorderSideReport {}

  /** Reports a visible border whose color is Excel's implicit default. */
  @Variant(
      id = "DEFAULT_COLOR",
      summary = "Read one visible border side using Excel's implicit default color.")
  record DefaultColor(ExcelBorderStyle style) implements CellBorderSideReport {
    public DefaultColor {
      requireVisibleStyle(style);
    }
  }

  /** Reports a visible border with one explicit color reference. */
  @Variant(
      id = "COLORED",
      summary = "Read one visible border side with an explicit color reference.")
  record Colored(ExcelBorderStyle style, CellColorReport color) implements CellBorderSideReport {
    public Colored {
      requireVisibleStyle(style);
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  private static void requireVisibleStyle(ExcelBorderStyle style) {
    Objects.requireNonNull(style, "style must not be null");
    if (style == ExcelBorderStyle.NONE) {
      throw new IllegalArgumentException("visible border report style must not be NONE");
    }
  }
}
