package dev.erst.gridgrind.core.model.formula;

import java.util.Objects;

/** Signals invalid OOXML formula character data supplied through the opaque formula input. */
public final class InvalidRawFormulaTextException extends IllegalArgumentException
    implements FormulaInputException {
  private static final long serialVersionUID = 1L;
  private final String publicMessage;
  private final dev.erst.gridgrind.core.problem.PayloadLocation fieldLocation;

  /** Creates one failure with a public, request-safe explanation. */
  public InvalidRawFormulaTextException(String message) {
    super(Objects.requireNonNull(message, "message must not be null"));
    this.publicMessage = message;
    this.fieldLocation = new dev.erst.gridgrind.core.problem.PayloadLocation.Unavailable();
  }

  /** Creates one failure while retaining the lower-level validation cause. */
  public InvalidRawFormulaTextException(String message, Throwable cause) {
    super(Objects.requireNonNull(message, "message must not be null"), cause);
    this.publicMessage = message;
    this.fieldLocation = new dev.erst.gridgrind.core.problem.PayloadLocation.Unavailable();
  }

  /** Creates a classified failure at the caller-owned relative field path. */
  public InvalidRawFormulaTextException(String message, String fieldPath) {
    super(Objects.requireNonNull(message, "message must not be null"));
    this.publicMessage = message;
    this.fieldLocation =
        new dev.erst.gridgrind.core.problem.PayloadLocation.PathOnly(
            Objects.requireNonNull(fieldPath, "fieldPath must not be null"));
  }

  /** Retains the classified location and lower-level validation cause. */
  public InvalidRawFormulaTextException(String message, String fieldPath, Throwable cause) {
    super(Objects.requireNonNull(message, "message must not be null"), cause);
    this.publicMessage = message;
    this.fieldLocation = new dev.erst.gridgrind.core.problem.PayloadLocation.PathOnly(fieldPath);
  }

  @Override
  public java.util.Optional<String> fieldPath() {
    return fieldLocation.jsonPath();
  }

  @Override
  public String publicMessage() {
    return publicMessage;
  }
}
