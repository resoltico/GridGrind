package dev.erst.gridgrind.core.model.value;

/** Factual reason why an observed temporal format has no supported civil interpretation. */
public enum CellTemporalReason {
  NONFINITE,
  NEGATIVE,
  OUTSIDE_CIVIL_RANGE,
  FICTITIOUS_1900_LEAP_DAY,
  UNSUPPORTED_ELAPSED_FORMAT
}
