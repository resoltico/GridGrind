package dev.erst.gridgrind.executor;

/** Failure raised when a file-backed authored input path does not exist. */
final class InputSourceNotFoundException extends InputSourceException {
  private static final long serialVersionUID = 1L;

  InputSourceNotFoundException(
      String message, String inputKind, String inputPath, Throwable cause) {
    super(message, inputKind, inputPath, cause);
  }
}
