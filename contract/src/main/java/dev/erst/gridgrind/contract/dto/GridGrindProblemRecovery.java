package dev.erst.gridgrind.contract.dto;

/** Recommended next action for the caller after receiving a problem. */
public enum GridGrindProblemRecovery {
  CHANGE_REQUEST,
  CHECK_ENVIRONMENT,
  ESCALATE
}
