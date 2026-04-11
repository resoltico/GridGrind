package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;

/** Test-only formula runtime that replaces reflection-heavy evaluator proxies with direct seams. */
final class FormulaRuntimeTestDouble implements ExcelFormulaRuntime {
  private final ExcelFormulaRuntime delegate;
  private final RuntimeException evaluateFailure;
  private final RuntimeException displayFailure;
  private final boolean returnNullEvaluation;

  private FormulaRuntimeTestDouble(
      ExcelFormulaRuntime delegate,
      RuntimeException evaluateFailure,
      RuntimeException displayFailure,
      boolean returnNullEvaluation) {
    this.delegate = delegate;
    this.evaluateFailure = evaluateFailure;
    this.displayFailure = displayFailure;
    this.returnNullEvaluation = returnNullEvaluation;
  }

  /** Returns a runtime that delegates all behavior to the supplied POI evaluator. */
  static FormulaRuntimeTestDouble delegating(FormulaEvaluator formulaEvaluator) {
    return new FormulaRuntimeTestDouble(
        ExcelFormulaRuntime.poi(formulaEvaluator), null, null, false);
  }

  /** Returns a runtime whose evaluate() path throws while display rendering still delegates. */
  static FormulaRuntimeTestDouble failingEvaluation(
      FormulaEvaluator formulaEvaluator, RuntimeException evaluateFailure) {
    return new FormulaRuntimeTestDouble(
        ExcelFormulaRuntime.poi(formulaEvaluator),
        Objects.requireNonNull(evaluateFailure, "evaluateFailure must not be null"),
        null,
        false);
  }

  /** Returns a runtime whose displayValue() path throws while evaluate() still delegates. */
  static FormulaRuntimeTestDouble failingDisplay(
      FormulaEvaluator formulaEvaluator, RuntimeException displayFailure) {
    return new FormulaRuntimeTestDouble(
        ExcelFormulaRuntime.poi(formulaEvaluator),
        null,
        Objects.requireNonNull(displayFailure, "displayFailure must not be null"),
        false);
  }

  /**
   * Returns a runtime whose evaluate() path yields null while display rendering still delegates.
   */
  static FormulaRuntimeTestDouble nullEvaluation(FormulaEvaluator formulaEvaluator) {
    return new FormulaRuntimeTestDouble(
        ExcelFormulaRuntime.poi(formulaEvaluator), null, null, true);
  }

  /** Returns a runtime that throws for both evaluation and display rendering. */
  static FormulaRuntimeTestDouble alwaysFail(RuntimeException exception) {
    RuntimeException failure = Objects.requireNonNull(exception, "exception must not be null");
    return new FormulaRuntimeTestDouble(null, failure, failure, false);
  }

  @Override
  public CellValue evaluate(Cell cell) {
    Objects.requireNonNull(cell, "cell must not be null");
    if (evaluateFailure != null) {
      throw evaluateFailure;
    }
    if (returnNullEvaluation) {
      return null;
    }
    return Objects.requireNonNull(delegate, "delegate must not be null").evaluate(cell);
  }

  @Override
  public CellType evaluateFormulaCell(Cell cell) {
    Objects.requireNonNull(cell, "cell must not be null");
    if (evaluateFailure != null) {
      throw evaluateFailure;
    }
    if (returnNullEvaluation) {
      return CellType._NONE;
    }
    return Objects.requireNonNull(delegate, "delegate must not be null").evaluateFormulaCell(cell);
  }

  @Override
  public void clearCachedResults() {
    if (delegate != null) {
      delegate.clearCachedResults();
    }
  }

  @Override
  public String displayValue(DataFormatter formatter, Cell cell) {
    Objects.requireNonNull(formatter, "formatter must not be null");
    Objects.requireNonNull(cell, "cell must not be null");
    if (displayFailure != null) {
      throw displayFailure;
    }
    return Objects.requireNonNull(delegate, "delegate must not be null")
        .displayValue(formatter, cell);
  }

  @Override
  public ExcelFormulaRuntimeContext context() {
    return delegate == null
        ? ExcelFormulaEnvironment.defaults().runtimeContext()
        : delegate.context();
  }
}
