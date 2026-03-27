package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.CellReference;

/** Immutable workbook-core definition of one named range to create or replace. */
public record ExcelNamedRangeDefinition(
    String name, ExcelNamedRangeScope scope, ExcelNamedRangeTarget target) {
  public ExcelNamedRangeDefinition {
    name = validateName(name);
    Objects.requireNonNull(scope, "scope must not be null");
    Objects.requireNonNull(target, "target must not be null");
  }

  /** Validates and canonicalizes one defined-name identifier. */
  public static String validateName(String name) {
    Objects.requireNonNull(name, "name must not be null");
    if (name.isBlank()) {
      throw new IllegalArgumentException("name must not be blank");
    }
    if (!name.matches("^[A-Za-z_][A-Za-z0-9_.]*$")) {
      throw new IllegalArgumentException(
          "name must start with a letter or underscore and contain only letters, digits, underscore, or period");
    }
    if (name.startsWith("_xlnm.") || name.startsWith("_XLNM.")) {
      throw new IllegalArgumentException("name must not use the reserved _xlnm. prefix");
    }
    if (CellReference.classifyCellReference(name, SpreadsheetVersion.EXCEL2007)
        == CellReference.NameType.CELL) {
      throw new IllegalArgumentException(
          "name must not collide with A1-style cell reference syntax");
    }
    if (name.matches("(?i)^R[1-9][0-9]*C[1-9][0-9]*$")) {
      throw new IllegalArgumentException(
          "name must not collide with R1C1-style cell reference syntax");
    }
    return name;
  }
}
