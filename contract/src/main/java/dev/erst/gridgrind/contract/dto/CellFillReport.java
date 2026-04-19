package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.ExcelFillPattern;
import java.util.Objects;

/** Effective fill facts reported with every analyzed cell. */
public record CellFillReport(
    ExcelFillPattern pattern,
    CellColorReport foregroundColor,
    CellColorReport backgroundColor,
    CellGradientFillReport gradient) {
  /** Creates a non-gradient fill report. */
  public CellFillReport(
      ExcelFillPattern pattern, CellColorReport foregroundColor, CellColorReport backgroundColor) {
    this(pattern, foregroundColor, backgroundColor, null);
  }

  public CellFillReport {
    Objects.requireNonNull(pattern, "pattern must not be null");
    if (gradient != null) {
      if (foregroundColor != null || backgroundColor != null) {
        throw new IllegalArgumentException(
            "gradient fills must not also expose foregroundColor or backgroundColor");
      }
    } else {
      if (pattern == ExcelFillPattern.NONE
          && (foregroundColor != null || backgroundColor != null)) {
        throw new IllegalArgumentException("fill pattern NONE does not accept colors");
      }
      if (pattern == ExcelFillPattern.SOLID && backgroundColor != null) {
        throw new IllegalArgumentException("fill backgroundColor is not supported for SOLID fills");
      }
    }
  }
}
