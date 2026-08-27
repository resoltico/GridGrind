package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Authored formula facts that never require evaluator execution. */
record ExcelFormulaCell(String address, String formula) {
  ExcelFormulaCell {
    Objects.requireNonNull(address, "address must not be null");
    Objects.requireNonNull(formula, "formula must not be null");
  }
}
