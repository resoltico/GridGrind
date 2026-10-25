package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.result.WorkbookResultSupport;
import java.util.List;
import java.util.Objects;

/** One grouped formula pattern and the addresses where it appears. */
public record FormulaPatternReport(String formula, int occurrenceCount, List<String> addresses) {
  public FormulaPatternReport {
    Objects.requireNonNull(formula, "formula must not be null");
    if (formula.isBlank()) {
      throw new IllegalArgumentException("formula must not be blank");
    }
    if (occurrenceCount <= 0) {
      throw new IllegalArgumentException("occurrenceCount must be greater than 0");
    }
    addresses = WorkbookResultSupport.copyStrings(addresses, "addresses");
  }
}
