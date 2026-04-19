package dev.erst.gridgrind.executor;

/** Failure raised when one authored input cannot be loaded from disk. */
final class InputSourceReadException extends InputSourceException {
  private static final long serialVersionUID = 1L;

  InputSourceReadException(String message, String inputKind, String inputPath, Throwable cause) {
    super(message, inputKind, inputPath, cause);
  }
}
