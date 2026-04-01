package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Supported table-style definition for GridGrind-authored tables. */
public sealed interface ExcelTableStyle permits ExcelTableStyle.None, ExcelTableStyle.Named {

  /** Clear table-style metadata and leave the table unstyled. */
  record None() implements ExcelTableStyle {}

  /** Apply one named workbook table style with explicit stripe and emphasis flags. */
  record Named(
      String name,
      boolean showFirstColumn,
      boolean showLastColumn,
      boolean showRowStripes,
      boolean showColumnStripes)
      implements ExcelTableStyle {
    public Named {
      Objects.requireNonNull(name, "name must not be null");
      if (name.isBlank()) {
        throw new IllegalArgumentException("name must not be blank");
      }
    }
  }
}
