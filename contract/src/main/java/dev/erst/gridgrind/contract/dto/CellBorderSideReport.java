package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import java.util.Objects;

/** Effective facts for one resolved cell-border side. */
public record CellBorderSideReport(ExcelBorderStyle style, CellColorReport color) {
  public CellBorderSideReport {
    Objects.requireNonNull(style, "style must not be null");
  }
}
