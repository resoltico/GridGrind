package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable snapshot of the resolved border currently applied to a cell. */
public record ExcelBorderSnapshot(
    ExcelBorderSideSnapshot top,
    ExcelBorderSideSnapshot right,
    ExcelBorderSideSnapshot bottom,
    ExcelBorderSideSnapshot left) {
  public ExcelBorderSnapshot {
    Objects.requireNonNull(top, "top must not be null");
    Objects.requireNonNull(right, "right must not be null");
    Objects.requireNonNull(bottom, "bottom must not be null");
    Objects.requireNonNull(left, "left must not be null");
  }
}
