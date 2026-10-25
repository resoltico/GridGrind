package dev.erst.gridgrind.contract.dto;

/** High-level problem groupings so agents can route failures by class of issue. */
public enum GridGrindProblemCategory {
  ARGUMENTS,
  REQUEST,
  ASSERTION,
  FORMULA,
  RESOURCE,
  SECURITY,
  IO,
  INTERNAL
}
