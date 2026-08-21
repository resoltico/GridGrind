package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Request-scoped collector for independently resolvable authored-input failures. */
final class InputResolutionFailures {
  private final List<InputResolutionFailure> failures = new ArrayList<>();

  void add(Exception exception, Optional<JsonLocation> json, Object source) {
    failures.add(
        InputResolutionFailure.forSource(
            Objects.requireNonNull(exception, "exception must not be null"),
            Objects.requireNonNullElseGet(json, Optional::empty),
            source));
  }

  void throwIfAny() throws InputResolutionBatchException {
    if (!failures.isEmpty()) {
      throw new InputResolutionBatchException(failures);
    }
  }
}
