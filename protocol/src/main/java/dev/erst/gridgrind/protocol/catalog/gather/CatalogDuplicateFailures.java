package dev.erst.gridgrind.protocol.catalog.gather;

import java.util.Objects;

/** Builds protocol-catalog duplicate failures with one canonical message shape. */
public final class CatalogDuplicateFailures {
  private CatalogDuplicateFailures() {}

  /** Returns the canonical duplicate-entry failure for protocol-catalog construction. */
  public static IllegalStateException duplicateEntryFailure(
      String label, Object left, Object right) {
    Objects.requireNonNull(label, "label must not be null");
    if (label.isBlank()) {
      throw new IllegalArgumentException("label must not be blank");
    }
    return new IllegalStateException(
        "Duplicate %s detected while building the protocol catalog: %s / %s"
            .formatted(label, left, right));
  }
}
