package dev.erst.gridgrind.core.model.formula;

/** Outcome status for the explicit calculation phase of one request. */
public enum CalculationExecutionStatus {
  NOT_REQUESTED,
  SKIPPED,
  PARTIAL,
  SUCCEEDED,
  FAILED
}
