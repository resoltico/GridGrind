package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Product-owned validation and normalization rules for persisted pivot-table names. */
public final class ExcelPivotTableNaming {
  private ExcelPivotTableNaming() {}

  /** Validates and returns one pivot-table name exactly as authored. */
  public static String validateName(String name) {
    Objects.requireNonNull(name, "name must not be null");
    if (name.isBlank()) {
      throw new IllegalArgumentException("name must not be blank");
    }
    return name;
  }
}
