package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;

/** Factual array-formula group metadata returned by workbook inspection. */
@CatalogType(
    id = "ArrayFormulaReport",
    summary = "One factual array-formula group report returned by GET_ARRAY_FORMULAS.",
    displayOrder = 100)
public record ArrayFormulaReport(
    SheetLabel sheetName,
    RangeLabel range,
    CellAddress topLeftAddress,
    String formula,
    boolean singleCell) {
  public ArrayFormulaReport {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    java.util.Objects.requireNonNull(range, "range must not be null");
    java.util.Objects.requireNonNull(topLeftAddress, "topLeftAddress must not be null");
    formula = TextValidation.requireNonBlank(formula, "formula");
  }
}
