package dev.erst.gridgrind.contract.json;

/** Signals that a protocol payload could not be parsed as valid JSON. */
public final class InvalidJsonException extends IllegalArgumentException
    implements PayloadException {
  private static final long serialVersionUID = 1L;

  private final String jsonPath;
  private final Integer jsonLine;
  private final Integer jsonColumn;

  /** Creates the exception with the given message, JSON location, and parse cause. */
  public InvalidJsonException(
      String message, String jsonPath, Integer jsonLine, Integer jsonColumn, Throwable cause) {
    super(message, cause);
    this.jsonPath = jsonPath;
    this.jsonLine = jsonLine;
    this.jsonColumn = jsonColumn;
  }

  @Override
  public String jsonPath() {
    return jsonPath;
  }

  @Override
  public Integer jsonLine() {
    return jsonLine;
  }

  @Override
  public Integer jsonColumn() {
    return jsonColumn;
  }
}
