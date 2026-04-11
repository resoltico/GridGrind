package dev.erst.gridgrind.excel;

import java.nio.file.Path;
import java.util.Objects;

/** One external workbook binding available to the formula evaluator. */
public record ExcelFormulaExternalWorkbookBinding(String workbookName, Path path) {
  public ExcelFormulaExternalWorkbookBinding {
    Objects.requireNonNull(workbookName, "workbookName must not be null");
    if (workbookName.isBlank()) {
      throw new IllegalArgumentException("workbookName must not be blank");
    }
    Objects.requireNonNull(path, "path must not be null");
  }
}
