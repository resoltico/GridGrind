package dev.erst.gridgrind.engine.runtime;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Request-scoped collector for independently resolvable authored-input failures. */
final class InputResolutionFailures {
  private final List<Exception> failures = new ArrayList<>();

  void add(Exception failure) {
    failures.add(Objects.requireNonNull(failure, "failure must not be null"));
  }

  void throwIfAny() throws InputResolutionBatchException {
    if (!failures.isEmpty()) {
      throw new InputResolutionBatchException(failures);
    }
  }
}
