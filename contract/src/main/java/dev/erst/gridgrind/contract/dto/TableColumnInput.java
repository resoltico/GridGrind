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
    uniqueName = normalize(uniqueName);
    totalsRowLabel = normalize(totalsRowLabel);
    totalsRowFunction = normalizeTotalsRowFunction(totalsRowFunction);
    calculatedColumnFormula = normalize(calculatedColumnFormula);
  }

  private static String normalize(String value) {
    if (value == null) {
      return "";
    }
    return value;
  }

  private static String normalizeTotalsRowFunction(String value) {
    if (value == null) {
      return "";
    }
    String normalized = value.trim();
    return normalized.isEmpty() ? "" : normalized.toLowerCase(Locale.ROOT);
  }
}
