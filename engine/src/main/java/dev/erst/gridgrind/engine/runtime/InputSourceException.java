package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.io.IOException;
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
    this.inputKind = TextValidation.requireNonBlank(inputKind, "inputKind");
    this.inputPath = inputPath;
  }

  String inputKind() {
    return inputKind;
  }

  @Nullable String inputPath() {
    return inputPath;
  }
}
