package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.result.WorkbookResultSupport;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;

/** One grouped formula pattern and the addresses where it appears. */
public record FormulaPatternReport(
    String formula, int occurrenceCount, List<CellAddress> addresses) {
  public FormulaPatternReport {
    formula = TextValidation.requireNonBlank(formula, "formula");
    if (occurrenceCount <= 0) {
      throw new IllegalArgumentException("occurrenceCount must be greater than 0");
    }
    addresses = WorkbookResultSupport.copyValues(addresses, "addresses");
  }
}
