package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable workbook-core definition of one table to create or replace. */
public record ExcelTableDefinition(
    String name, String sheetName, String range, boolean showTotalsRow, ExcelTableStyle style) {
  public ExcelTableDefinition {
    name = validateName(name);
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    if (sheetName.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }
    if (sheetName.length() > 31) {
      throw new IllegalArgumentException("sheetName must not exceed 31 characters: " + sheetName);
    }
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
