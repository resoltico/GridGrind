package dev.erst.gridgrind.engine.runtime;

import java.util.Objects;

/** A caller-owned resource bound in addition to the host's aggregate snapshot allowance. */
record InputSnapshotLimit(long maxBytes, RuntimeException refusal) {
  InputSnapshotLimit {
    if (maxBytes < 0)
      throw new IllegalArgumentException("Snapshot resource bound must not be negative");
    Objects.requireNonNull(refusal, "refusal");
  }

  void require(long bytes) {
    if (bytes > maxBytes) throw refusal;
  }
}
