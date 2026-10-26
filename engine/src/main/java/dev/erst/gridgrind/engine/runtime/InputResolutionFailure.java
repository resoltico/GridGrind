package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.InputReference;
import dev.erst.gridgrind.core.source.BinarySourceSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.util.Objects;
import java.util.Optional;

/** One source-resolution failure paired with the authored input field that produced it. */
record InputResolutionFailure(
    Exception exception, Optional<JsonLocation> json, InputReference input) {
  InputResolutionFailure {
    Objects.requireNonNull(exception, "exception must not be null");
    json = Objects.requireNonNullElseGet(json, Optional::empty);
    Objects.requireNonNull(input, "input must not be null");
  }

  static InputResolutionFailure unlocated(Exception exception) {
    return new InputResolutionFailure(exception, Optional.empty(), InputReference.unknown());
  }

  static InputResolutionFailure forSource(
      Exception exception, Optional<JsonLocation> json, Object source) {
    return new InputResolutionFailure(exception, json, inputReference(source));
  }

  private static InputReference inputReference(Object source) {
    return switch (source) {
      case TextSourceSpec.Utf8File file -> InputReference.path("source-backed text", file.path());
      case TextSourceSpec.Stdin _ -> InputReference.kind("source-backed text");
      case BinarySourceSpec.File file -> InputReference.path("source-backed binary", file.path());
      case BinarySourceSpec.Stdin _ -> InputReference.kind("source-backed binary");
      case TextSourceSpec.Inline _ -> InputReference.unknown();
      case BinarySourceSpec.InlineBase64 _ -> InputReference.unknown();
      case null -> InputReference.unknown();
      default -> InputReference.unknown();
    };
  }
}
