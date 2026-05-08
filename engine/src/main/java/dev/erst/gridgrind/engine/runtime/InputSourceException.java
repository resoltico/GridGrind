package dev.erst.gridgrind.engine.runtime;

import java.io.IOException;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Base exception for failures while resolving source-backed authored inputs. */
abstract sealed class InputSourceException extends IOException
    permits InputSourceNotFoundException,
        InputSourceReadException,
        InputSourceUnavailableException {
  private static final long serialVersionUID = 1L;

  private final String inputKind;
  private final @Nullable String inputPath;

  InputSourceException(
      String message, String inputKind, @Nullable String inputPath, @Nullable Throwable cause) {
    super(message, cause);
    this.inputKind = requireNonBlank(inputKind, "inputKind");
    this.inputPath = inputPath;
  }

  String inputKind() {
    return inputKind;
  }

  @Nullable String inputPath() {
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
