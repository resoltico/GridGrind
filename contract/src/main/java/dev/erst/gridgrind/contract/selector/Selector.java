package dev.erst.gridgrind.contract.selector;

import tools.jackson.databind.annotation.JsonSerialize;

/** Marker for one immutable workbook-target selector family. */
@JsonSerialize(using = SelectorJsonSerializer.class)
@FunctionalInterface
public interface Selector {
  /** Returns the selector's declared target-count contract. */
  SelectorCardinality cardinality();
}
