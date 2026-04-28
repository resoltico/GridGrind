package dev.erst.gridgrind.executor;

import java.util.Objects;
import java.util.Optional;

/** Identity-preserving helpers for source-backed authored payload resolution. */
final class SourceBackedResolutionIdentitySupport {
  private SourceBackedResolutionIdentitySupport() {}

  /**
   * Preserves authored step and selector object identity when source-backed resolution makes no
   * semantic change, which avoids rebuilding canonical records unnecessarily.
   */
  @SuppressWarnings("PMD.CompareObjectsWithEquals")
  static boolean sameReference(Object left, Object right) {
    return left == right;
  }

  static <T> boolean sameOptionalReference(Optional<T> left, Optional<T> right) {
    Objects.requireNonNull(left, "left must not be null");
    Objects.requireNonNull(right, "right must not be null");
    return left.isPresent() == right.isPresent()
        && (left.isEmpty() || sameReference(left.orElseThrow(), right.orElseThrow()));
  }
}
