package dev.erst.gridgrind.protocol.dto;

/** High-level problem groupings so agents can route failures by class of issue. */
public enum GridGrindProblemCategory {
  ARGUMENTS,
  REQUEST,
  FORMULA,
  RESOURCE,
  SECURITY,
  IO,
  INTERNAL
}
