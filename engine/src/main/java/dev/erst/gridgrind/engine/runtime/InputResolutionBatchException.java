package dev.erst.gridgrind.engine.runtime;

import java.io.IOException;
import java.util.List;
import java.util.Objects;

/** Aggregate raised after a source-resolution pass has collected every independent failure. */
final class InputResolutionBatchException extends IOException {
  private static final long serialVersionUID = 1L;

  private final List<InputResolutionFailure> failures;

  InputResolutionBatchException(List<InputResolutionFailure> failures) {
    super("Unable to resolve one or more authored inputs");
    this.failures = List.copyOf(Objects.requireNonNull(failures, "failures must not be null"));
  }

  List<InputResolutionFailure> failures() {
    return failures;
  }
}
