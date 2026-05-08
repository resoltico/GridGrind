package dev.erst.gridgrind.engine.runtime;

/** Failure raised when one authored input requests STANDARD_INPUT but none is bound. */
final class InputSourceUnavailableException extends InputSourceException {
  private static final long serialVersionUID = 1L;

  InputSourceUnavailableException(String message, String inputKind) {
    super(message, inputKind, null, null);
  }
}
