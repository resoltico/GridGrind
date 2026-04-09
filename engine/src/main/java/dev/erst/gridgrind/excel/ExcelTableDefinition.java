package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable workbook-core definition of one table to create or replace. */
public record ExcelTableDefinition(
    String name, String sheetName, String range, boolean showTotalsRow, ExcelTableStyle style) {
  public ExcelTableDefinition {
    name = validateName(name);
    ExcelSheetNames.requireValid(sheetName, "sheetName");
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    Objects.requireNonNull(style, "style must not be null");
  }

  /** Validates and canonicalizes one table identifier. */
  public static String validateName(String name) {
    return ExcelNamedRangeDefinition.validateName(name);
  }
}
