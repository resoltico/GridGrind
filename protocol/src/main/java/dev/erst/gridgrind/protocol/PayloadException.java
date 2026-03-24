package dev.erst.gridgrind.protocol;

/** Base class for JSON payload failures with structured path and location metadata. */
public abstract sealed class PayloadException extends IllegalArgumentException
    permits InvalidJsonException, InvalidRequestException {
  private static final long serialVersionUID = 1L;

  private final String jsonPath;
  private final Integer jsonLine;
  private final Integer jsonColumn;

  protected PayloadException(
      String message, String jsonPath, Integer jsonLine, Integer jsonColumn, Throwable cause) {
    super(message, cause);
    this.jsonPath = jsonPath;
    this.jsonLine = jsonLine;
    this.jsonColumn = jsonColumn;
  }

  public String jsonPath() {
    return jsonPath;
  }

  public Integer jsonLine() {
    return jsonLine;
  }

  public Integer jsonColumn() {
    return jsonColumn;
  }
}
