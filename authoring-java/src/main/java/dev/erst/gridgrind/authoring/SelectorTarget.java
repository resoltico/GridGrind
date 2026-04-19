package dev.erst.gridgrind.authoring;

import dev.erst.gridgrind.contract.selector.Selector;

/** Public seam for fluent authoring targets that compile to canonical selectors. */
@FunctionalInterface
public interface SelectorTarget {
  /** Returns the canonical selector emitted by this target. */
  Selector selector();
}
