package dev.erst.gridgrind.contract.dto;

import java.util.Locale;
import java.util.Objects;

/** Protocol-owned validation helpers for defined names and related identifiers. */
public final class ProtocolDefinedNameValidation {
  private ProtocolDefinedNameValidation() {}

  /** Validates one protocol-facing defined-name identifier and returns its canonical text. */
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
    if (looksLikeA1CellReference(name)) {
      throw new IllegalArgumentException(
          "name must not collide with A1-style cell reference syntax");
    }
    if (name.matches("(?i)^R[1-9][0-9]*C[1-9][0-9]*$")) {
      throw new IllegalArgumentException(
          "name must not collide with R1C1-style cell reference syntax");
    }
    return name;
  }

  private static boolean looksLikeA1CellReference(String candidate) {
    if (!candidate.matches("(?i)^\\$?[A-Z]{1,3}\\$?[1-9][0-9]*$")) {
      return false;
    }
    String normalized = candidate.replace("$", "").toUpperCase(Locale.ROOT);
    String columnLabel = normalized.replaceAll("\\d.*$", "");
    int columnNumber = 0;
    for (int index = 0; index < columnLabel.length(); index++) {
      columnNumber = (columnNumber * 26) + (columnLabel.charAt(index) - 'A' + 1);
    }
    return columnNumber <= 16384;
  }
}
