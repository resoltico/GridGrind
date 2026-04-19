package dev.erst.gridgrind.contract.dto;

/** Outcome status for the explicit calculation phase of one request. */
public enum CalculationExecutionStatus {
  NOT_REQUESTED,
  SKIPPED,
  SUCCEEDED,
  FAILED
}
