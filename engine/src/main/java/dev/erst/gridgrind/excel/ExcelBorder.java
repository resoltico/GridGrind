package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.Optional;

/**
 * Border patch applied through {@link ExcelCellStyle}, with optional defaults and side overrides.
 */
public record ExcelBorder(
    Optional<ExcelBorderSide> all,
    Optional<ExcelBorderSide> top,
    Optional<ExcelBorderSide> right,
    Optional<ExcelBorderSide> bottom,
    Optional<ExcelBorderSide> left) {
  public ExcelBorder {
    Objects.requireNonNull(all, "all must not be null");
    Objects.requireNonNull(top, "top must not be null");
    Objects.requireNonNull(right, "right must not be null");
    Objects.requireNonNull(bottom, "bottom must not be null");
    Objects.requireNonNull(left, "left must not be null");
    if (all.isEmpty() && top.isEmpty() && right.isEmpty() && bottom.isEmpty() && left.isEmpty()) {
      throw new IllegalArgumentException("border must set at least one side");
    }
  }
}
