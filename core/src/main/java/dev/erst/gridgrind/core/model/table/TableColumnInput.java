package dev.erst.gridgrind.core.model.table;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.WireDefault;
import java.util.Locale;

/** Advanced authored table-column metadata applied by ordinal column index. */
@CatalogType(
    id = "TableColumnInput",
    summary = "Advanced table-column metadata applied by zero-based ordinal column index.",
    displayOrder = 73)
public record TableColumnInput(
    int columnIndex,
    @WireDefault String uniqueName,
    @WireDefault String totalsRowLabel,
    @WireDefault String totalsRowFunction,
    @WireDefault String calculatedColumnFormula) {

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
