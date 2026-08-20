package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import java.util.Objects;
import java.util.Optional;

/** One source-resolution failure paired with the authored input field that produced it. */
record InputResolutionFailure(Exception exception, Optional<JsonLocation> json) {
  InputResolutionFailure {
    Objects.requireNonNull(exception, "exception must not be null");
    json = Objects.requireNonNullElseGet(json, Optional::empty);
  }

  static InputResolutionFailure unlocated(Exception exception) {
    return new InputResolutionFailure(exception, Optional.empty());
  }
}
