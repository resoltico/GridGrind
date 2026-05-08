package dev.erst.gridgrind.engine.runtime;

import org.jspecify.annotations.Nullable;

/** Failure raised when one authored input cannot be loaded from disk. */
final class InputSourceReadException extends InputSourceException {
  private static final long serialVersionUID = 1L;

  InputSourceReadException(
      String message, String inputKind, @Nullable String inputPath, @Nullable Throwable cause) {
    super(message, inputKind, inputPath, cause);
  }
}
