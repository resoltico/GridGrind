package dev.erst.gridgrind.protocol;

/** Signals that a parsed payload violated GridGrind request validation rules. */
public final class InvalidRequestException extends PayloadException {
  private static final long serialVersionUID = 1L;

  /** Creates the exception with the given message, JSON location, and validation cause. */
  public InvalidRequestException(
      String message, String jsonPath, Integer jsonLine, Integer jsonColumn, Throwable cause) {
    super(message, jsonPath, jsonLine, jsonColumn, cause);
  }
}
