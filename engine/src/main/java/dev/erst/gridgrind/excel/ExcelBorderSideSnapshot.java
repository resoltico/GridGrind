package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** One resolved workbook border side preserving style and factual color semantics. */
public record ExcelBorderSideSnapshot(ExcelBorderStyle style, @Nullable ExcelColorSnapshot color) {
  public ExcelBorderSideSnapshot {
    Objects.requireNonNull(style, "style must not be null");
  }
}
