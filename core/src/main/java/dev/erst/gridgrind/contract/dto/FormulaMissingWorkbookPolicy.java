package dev.erst.gridgrind.contract.dto;

/** Controls how formula evaluation handles missing external workbook references. */
public enum FormulaMissingWorkbookPolicy {
  /** Reject evaluation when any referenced external workbook is missing or unbound. */
  ERROR,

  /** Use the workbook's cached formula result when an external workbook cannot be resolved. */
  USE_CACHED_VALUE
}
