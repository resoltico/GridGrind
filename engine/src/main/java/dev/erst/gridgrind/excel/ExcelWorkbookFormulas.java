package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Formula evaluation, cache management, and capability inspection operations for one workbook. */
public final class ExcelWorkbookFormulas {
  private final ExcelWorkbook workbook;

  ExcelWorkbookFormulas(ExcelWorkbook workbook) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
  }

  /** Evaluates every formula cell currently present in the workbook. */
  public ExcelWorkbook evaluateAll() {
    return ExcelWorkbookFormulaSupport.evaluateAllFormulas(workbook);
  }

  /** Classifies every formula cell currently present in the workbook under the current runtime. */
  public List<ExcelFormulaCapabilityAssessment> assessAllCapabilities() {
    return ExcelWorkbookFormulaSupport.assessAllFormulaCapabilities(workbook);
  }

  /** Evaluates one or more explicit formula-cell targets and stores their cached results. */
  public ExcelWorkbook evaluate(List<ExcelFormulaCellTarget> cells) {
    return ExcelWorkbookFormulaSupport.evaluateFormulaCells(workbook, cells);
  }

  /** Classifies one explicit set of formula-cell targets under the current runtime. */
  public List<ExcelFormulaCapabilityAssessment> assessCapabilities(
      List<ExcelFormulaCellTarget> cells) {
    return ExcelWorkbookFormulaSupport.assessFormulaCellCapabilities(workbook, cells);
  }

  /** Clears persisted formula cached results and resets the in-process evaluator state. */
  public ExcelWorkbook clearCaches() {
    return ExcelWorkbookFormulaSupport.clearFormulaCaches(workbook);
  }

  /** Marks the workbook to recalculate formulas when opened in Excel-compatible clients. */
  public ExcelWorkbook markRecalculateOnOpen() {
    return ExcelWorkbookFormulaSupport.forceFormulaRecalculationOnOpen(workbook);
  }

  /** Returns whether the workbook is marked to recalculate formulas when opened in Excel. */
  public boolean recalculateOnOpenEnabled() {
    return workbook.xssfWorkbook().getForceFormulaRecalculation();
  }
}
