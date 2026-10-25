package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import java.util.Objects;

/** Factual array-formula group metadata returned by workbook inspection. */
public record ArrayFormulaReport(
    SheetLabel sheetName,
    String range,
    CellAddress topLeftAddress,
    String formula,
    boolean singleCell) {
  public ArrayFormulaReport {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    range = requireNonBlank(range, "range");
    java.util.Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
    formula = requireNonBlank(formula, "formula");
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
