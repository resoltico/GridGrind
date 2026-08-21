package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.selector.Selector;

/** Resolves selector families from a bound nested operation without a parallel schema. */
@FunctionalInterface
interface WorkbookOperationTargetSelectorDerivation {
  /** Derives the selector classes accepted by one bound nested operation. */
  Class<? extends Selector>[] acceptedSelectors(Object operation);
}
