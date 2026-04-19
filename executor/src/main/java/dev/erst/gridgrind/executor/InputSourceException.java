package dev.erst.gridgrind.executor;

import java.io.IOException;
import java.util.Objects;

/** Base exception for failures while resolving source-backed authored inputs. */
abstract sealed class InputSourceException extends IOException
    permits InputSourceNotFoundException,
        InputSourceReadException,
        InputSourceUnavailableException {
  private static final long serialVersionUID = 1L;

  private final String inputKind;
  private final String inputPath;

  InputSourceException(String message, String inputKind, String inputPath, Throwable cause) {
    super(message, cause);
    this.inputKind = requireNonBlank(inputKind, "inputKind");
    this.inputPath = inputPath;
  }

  String inputKind() {
    return inputKind;
  }

  String inputPath() {
    return inputPath;
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
