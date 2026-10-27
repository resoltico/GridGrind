package dev.erst.gridgrind.core.semantics;

/** Static confidence boundary for the workbook scope an operation can affect. */
public enum OperationEffectFootprint {
  TARGET_BOUNDED,
  CONSERVATIVE_WORKBOOK_WIDE
}
