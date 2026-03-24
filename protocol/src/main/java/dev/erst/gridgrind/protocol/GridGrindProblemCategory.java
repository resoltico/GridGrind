package dev.erst.gridgrind.protocol;

/** High-level problem groupings so agents can route failures by class of issue. */
public enum GridGrindProblemCategory {
  ARGUMENTS,
  REQUEST,
  FORMULA,
  RESOURCE,
  IO,
  INTERNAL
}
