package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable snapshot of the style currently applied to a cell. */
public record ExcelCellStyleSnapshot(
    String numberFormat,
    ExcelCellAlignmentSnapshot alignment,
    ExcelCellFontSnapshot font,
    ExcelCellFillSnapshot fill,
    ExcelBorderSnapshot border,
    ExcelCellProtectionSnapshot protection) {
  public ExcelCellStyleSnapshot {
    Objects.requireNonNull(numberFormat, "numberFormat must not be null");
    Objects.requireNonNull(alignment, "alignment must not be null");
    Objects.requireNonNull(font, "font must not be null");
    Objects.requireNonNull(fill, "fill must not be null");
    Objects.requireNonNull(border, "border must not be null");
    Objects.requireNonNull(protection, "protection must not be null");
  }
}
