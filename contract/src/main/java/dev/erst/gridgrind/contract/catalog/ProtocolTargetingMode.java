package dev.erst.gridgrind.contract.catalog;

/** Declares how one protocol leaf derives its accepted target-selector families. */
public enum ProtocolTargetingMode {
  /** The leaf owns one explicit static selector-family set. */
  STATIC,

  /** The leaf derives its selector families from one nested analysis query. */
  ANALYSIS_QUERY,

  /** The leaf derives its selector families from one nested assertion. */
  NESTED_ASSERTION,

  /** The leaf derives its selector families from the intersection of nested assertions. */
  INTERSECTION_OF_NESTED_ASSERTIONS
}
