package dev.erst.gridgrind.jazzer.tool;

/** Distinguishes active fuzzing runs from regression-only replay runs. */
public enum RunMode {
  ACTIVE_FUZZING,
  REGRESSION
}
