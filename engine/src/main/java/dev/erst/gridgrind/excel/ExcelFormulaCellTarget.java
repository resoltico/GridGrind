package dev.erst.gridgrind.excel;

import java.util.Objects;

/** One concrete formula cell to evaluate. */
public record ExcelFormulaCellTarget(String sheetName, String address) {
  public ExcelFormulaCellTarget {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    if (sheetName.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }
    Objects.requireNonNull(address, "address must not be null");
    if (address.isBlank()) {
      throw new IllegalArgumentException("address must not be blank");
    }
  }
}
