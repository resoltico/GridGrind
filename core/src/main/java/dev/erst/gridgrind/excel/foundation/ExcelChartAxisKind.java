package dev.erst.gridgrind.excel.foundation;

/** Supported chart-axis families. */
public enum ExcelChartAxisKind {
  CATEGORY,
  DATE,
  SERIES,
  VALUE;

  /** Returns true when the axis occupies the category/date family. */
  public boolean isCategoryAxis() {
    return this == CATEGORY || this == DATE;
  }

  /** Returns true when the axis is the series axis used by 3D chart families. */
  public boolean isSeriesAxis() {
    return this == SERIES;
  }

  /** Returns true when the axis is numeric/value-oriented. */
  public boolean isValueAxis() {
    return this == VALUE;
  }
}
