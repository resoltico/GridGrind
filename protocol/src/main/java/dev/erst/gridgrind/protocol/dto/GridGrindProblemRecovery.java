package dev.erst.gridgrind.protocol.dto;

/** Recommended next action for the caller after receiving a problem. */
public enum GridGrindProblemRecovery {
  CHANGE_REQUEST,
  CHECK_ENVIRONMENT,
  ESCALATE
}
