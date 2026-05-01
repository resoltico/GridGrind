package dev.erst.gridgrind.excel;

import java.util.Objects;

/** One factual sort condition stored inside an autofilter sort-state payload. */
public sealed interface ExcelAutofilterSortConditionSnapshot
    permits ExcelAutofilterSortConditionSnapshot.Value,
        ExcelAutofilterSortConditionSnapshot.CellColor,
        ExcelAutofilterSortConditionSnapshot.FontColor,
        ExcelAutofilterSortConditionSnapshot.Icon {
  /** Returns the persisted sheet range that this factual sort condition targets. */
  String range();

  /** Returns whether the persisted sort direction is descending. */
  boolean descending();

  /** One ordinary value-based factual sort condition. */
  record Value(String range, boolean descending) implements ExcelAutofilterSortConditionSnapshot {
    public Value {
      range = requireRange(range);
    }
  }

  /** One cell-fill-color factual sort condition. */
  record CellColor(String range, boolean descending, ExcelColorSnapshot color)
      implements ExcelAutofilterSortConditionSnapshot {
    public CellColor {
      range = requireRange(range);
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** One font-color factual sort condition. */
  record FontColor(String range, boolean descending, ExcelColorSnapshot color)
      implements ExcelAutofilterSortConditionSnapshot {
    public FontColor {
      range = requireRange(range);
      Objects.requireNonNull(color, "color must not be null");
    }
  }

  /** One icon factual sort condition. */
  record Icon(String range, boolean descending, int iconId)
      implements ExcelAutofilterSortConditionSnapshot {
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
