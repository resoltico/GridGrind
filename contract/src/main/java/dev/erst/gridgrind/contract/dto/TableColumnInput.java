package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Locale;

/** Advanced authored table-column metadata applied by ordinal column index. */
public record TableColumnInput(
    int columnIndex,
    String uniqueName,
    String totalsRowLabel,
    String totalsRowFunction,
    String calculatedColumnFormula) {
  /** Reads a table-column payload while defaulting omitted optional text fields to empty. */
  @JsonCreator
  public static TableColumnInput create(
      @JsonProperty("columnIndex") int columnIndex,
      @JsonProperty("uniqueName") String uniqueName,
      @JsonProperty("totalsRowLabel") String totalsRowLabel,
      @JsonProperty("totalsRowFunction") String totalsRowFunction,
      @JsonProperty("calculatedColumnFormula") String calculatedColumnFormula) {
    return new TableColumnInput(
        columnIndex,
        defaultEmpty(uniqueName),
        defaultEmpty(totalsRowLabel),
        defaultEmpty(totalsRowFunction),
        defaultEmpty(calculatedColumnFormula));
  }

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

  private static String defaultEmpty(String value) {
    return value == null ? "" : value;
  }
}
