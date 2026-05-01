package dev.erst.gridgrind.contract.dto;

import java.util.Locale;

/** Advanced authored table-column metadata applied by ordinal column index. */
public record TableColumnInput(
    int columnIndex,
    String uniqueName,
    String totalsRowLabel,
    String totalsRowFunction,
    String calculatedColumnFormula) {
  public TableColumnInput {
    if (columnIndex < 0) {
      throw new IllegalArgumentException("columnIndex must not be negative");
    }
    uniqueName = requireNonNull(uniqueName, "uniqueName");
    totalsRowLabel = requireNonNull(totalsRowLabel, "totalsRowLabel");
    totalsRowFunction =
        normalizeTotalsRowFunction(requireNonNull(totalsRowFunction, "totalsRowFunction"));
    calculatedColumnFormula = requireNonNull(calculatedColumnFormula, "calculatedColumnFormula");
  }

  private static String requireNonNull(String value, String fieldName) {
    return java.util.Objects.requireNonNull(value, fieldName + " must not be null");
  }

  private static String normalizeTotalsRowFunction(String value) {
    String normalized = value.trim();
    return normalized.isEmpty() ? "" : normalized.toLowerCase(Locale.ROOT);
  }
}
