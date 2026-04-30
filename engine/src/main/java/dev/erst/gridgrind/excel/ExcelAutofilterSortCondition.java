package dev.erst.gridgrind.excel;

import java.util.Objects;

/** One authored sort condition nested under a mutable-workbook autofilter sort state. */
public sealed interface ExcelAutofilterSortCondition
    permits ExcelAutofilterSortCondition.Value,
        ExcelAutofilterSortCondition.CellColor,
        ExcelAutofilterSortCondition.FontColor,
        ExcelAutofilterSortCondition.Icon {
  /** Returns the sheet range that this sort condition targets. */
  String range();

  /** Returns whether this sort condition orders the target range descending. */
  boolean descending();

  /** One ordinary value-based sort condition. */
  record Value(String range, boolean descending) implements ExcelAutofilterSortCondition {
    public Value {
      range = requireRange(range);
    }
  }

  /** One fill-color-based sort condition. */
  record CellColor(String range, boolean descending, ExcelColor color)
      implements ExcelAutofilterSortCondition {
    public CellColor {
      range = requireRange(range);
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** One font-color-based sort condition. */
  record FontColor(String range, boolean descending, ExcelColor color)
      implements ExcelAutofilterSortCondition {
    public FontColor {
      range = requireRange(range);
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** One icon-based sort condition. */
  record Icon(String range, boolean descending, int iconId)
      implements ExcelAutofilterSortCondition {
    public Icon {
      range = requireRange(range);
      if (iconId < 0) {
        throw new IllegalArgumentException("iconId must not be negative");
      }
    }
  }

  private static String requireRange(String range) {
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    return range;
  }
}
