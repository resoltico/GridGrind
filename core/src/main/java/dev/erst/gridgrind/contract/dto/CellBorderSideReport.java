package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import java.util.Objects;

/** Effective facts for one resolved cell-border side with explicit color semantics. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = CellBorderSideReport.None.class, name = "NONE"),
  @JsonSubTypes.Type(value = CellBorderSideReport.DefaultColor.class, name = "DEFAULT_COLOR"),
  @JsonSubTypes.Type(value = CellBorderSideReport.Colored.class, name = "COLORED")
})
public sealed interface CellBorderSideReport
    permits CellBorderSideReport.None,
        CellBorderSideReport.DefaultColor,
        CellBorderSideReport.Colored {
  /** Reports that this border side has no visible border. */
  record None() implements CellBorderSideReport {}

  /** Reports a visible border whose color is Excel's implicit default. */
  record DefaultColor(ExcelBorderStyle style) implements CellBorderSideReport {
    public DefaultColor {
      requireVisibleStyle(style);
    }
  }

  /** Reports a visible border with one explicit color reference. */
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
