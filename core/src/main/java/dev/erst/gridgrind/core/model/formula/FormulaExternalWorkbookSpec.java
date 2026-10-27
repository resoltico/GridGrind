package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** One external workbook binding used by formula evaluation. */
@CatalogType(
    id = "FormulaExternalWorkbookSpec",
    summary =
        "One external workbook binding keyed by the workbook name used inside formulas. Workbook paths resolve through the shared request-owned path rule.",
    noteRefs = {"requestOwnedPathRule"},
    displayOrder = 26)
public record FormulaExternalWorkbookSpec(String workbookName, String path) {
  public FormulaExternalWorkbookSpec {
    workbookName = requireWorkbookName(workbookName);
    WorkbookPlan.requireXlsxWorkbookPath(path);
  }

  private static String requireWorkbookName(String workbookName) {
    TextValidation.requireNonBlank(workbookName, "workbookName");
    if (workbookName.indexOf('[') >= 0 || workbookName.indexOf(']') >= 0) {
      throw new IllegalArgumentException(
          "workbookName must not contain [ or ] because formulas add brackets automatically");
    }
    return workbookName;
  }
}
