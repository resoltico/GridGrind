package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Factual table-style metadata loaded from a workbook. */
public sealed interface ExcelTableStyleSnapshot
    permits ExcelTableStyleSnapshot.None, ExcelTableStyleSnapshot.Named {

  /** Table has no style metadata. */
  record None() implements ExcelTableStyleSnapshot {}

  /** Table carries a raw named style plus explicit stripe and emphasis flags. */
  record Named(
      String name,
      boolean showFirstColumn,
      boolean showLastColumn,
      boolean showRowStripes,
      boolean showColumnStripes)
      implements ExcelTableStyleSnapshot {
    public Named {
      Objects.requireNonNull(name, "name must not be null");
    }
  }
}
