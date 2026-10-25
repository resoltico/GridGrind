package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** One external workbook binding used by formula evaluation. */
public record FormulaExternalWorkbookInput(String workbookName, String path) {
  public FormulaExternalWorkbookInput {
    workbookName = requireWorkbookName(workbookName);
    WorkbookPlan.requireXlsxWorkbookPath(path);
  }

  private static String requireWorkbookName(String workbookName) {
    Objects.requireNonNull(workbookName, "workbookName must not be null");
    if (workbookName.isBlank()) {
      throw new IllegalArgumentException("workbookName must not be blank");
    }
    if (workbookName.indexOf('[') >= 0 || workbookName.indexOf(']') >= 0) {
      throw new IllegalArgumentException(
          "workbookName must not contain [ or ] because formulas add brackets automatically");
    }
    return workbookName;
  }
}
