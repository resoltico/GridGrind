package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Effective facts for one resolved cell-border side. */
public record CellBorderSideReport(ExcelBorderStyle style, @Nullable CellColorReport color) {
  public CellBorderSideReport {
    Objects.requireNonNull(style, "style must not be null");
  }
}
