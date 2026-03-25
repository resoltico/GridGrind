package dev.erst.gridgrind.protocol;

/** Signals that a protocol payload could not be parsed as valid JSON. */
public final class InvalidJsonException extends PayloadException {
  private static final long serialVersionUID = 1L;

  /** Creates the exception with the given message, JSON location, and parse cause. */
  public InvalidJsonException(
      String message, String jsonPath, Integer jsonLine, Integer jsonColumn, Throwable cause) {
    super(message, jsonPath, jsonLine, jsonColumn, cause);
  }
}
