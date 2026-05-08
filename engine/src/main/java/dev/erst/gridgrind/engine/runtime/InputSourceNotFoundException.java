package dev.erst.gridgrind.engine.runtime;

import org.jspecify.annotations.Nullable;

/** Failure raised when a file-backed authored input path does not exist. */
final class InputSourceNotFoundException extends InputSourceException {
  private static final long serialVersionUID = 1L;

  InputSourceNotFoundException(
      String message, String inputKind, @Nullable String inputPath, @Nullable Throwable cause) {
    super(message, inputKind, inputPath, cause);
  }
}
