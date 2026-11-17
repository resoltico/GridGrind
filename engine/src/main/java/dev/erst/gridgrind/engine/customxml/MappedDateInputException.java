package dev.erst.gridgrind.engine.customxml;

import dev.erst.gridgrind.core.model.value.ExcelDateSystem;
import dev.erst.gridgrind.core.problem.ValueInvariantException;
import java.util.Objects;

/** A known mapped DATE refusal at an actual mapping occurrence, not a fabricated JSON cursor. */
public final class MappedDateInputException extends ValueInvariantException {
  private static final long serialVersionUID = 1L;
  private final ValueInvariantException coreProblem;
  private final String xpath;
  private final int mappingRowIndex;
  private final ExcelDateSystem dateSystem;

  /** Retains the core refusal and actual mapping occurrence that supplied its date operand. */
  public MappedDateInputException(
      ValueInvariantException problem, String xpath, int row, ExcelDateSystem dateSystem) {
    super(Objects.requireNonNull(problem).requestProblem());
    initCause(problem);
    this.coreProblem = problem;
    this.xpath = Objects.requireNonNull(xpath);
    this.mappingRowIndex = row;
    this.dateSystem = Objects.requireNonNull(dateSystem);
  }

  /** Returns the original core temporal refusal. */
  public ValueInvariantException coreProblem() {
    return coreProblem;
  }

  /** Returns the actual mapped field XPath. */
  public String xpath() {
    return xpath;
  }

  /** Returns the zero-based occurrence within the mapped field. */
  public int mappingRowIndex() {
    return mappingRowIndex;
  }

  /** Returns the receiving workbook calendar. */
  public ExcelDateSystem dateSystem() {
    return dateSystem;
  }
}
