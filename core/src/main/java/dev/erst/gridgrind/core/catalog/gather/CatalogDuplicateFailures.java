package dev.erst.gridgrind.core.catalog.gather;

import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** Builds protocol-catalog duplicate failures with one canonical message shape. */
public final class CatalogDuplicateFailures {
  private CatalogDuplicateFailures() {}

  /** Returns the canonical duplicate-entry failure for protocol-catalog construction. */
  public static IllegalStateException duplicateEntryFailure(
      String label, Object left, Object right) {
    TextValidation.requireNonBlank(label, "label");
    return new IllegalStateException(
        "Duplicate %s detected while building the protocol catalog: %s / %s"
            .formatted(label, left, right));
  }
}
