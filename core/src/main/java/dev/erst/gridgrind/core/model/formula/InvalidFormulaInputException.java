package dev.erst.gridgrind.core.model.formula;

import java.util.Objects;

/** Signals invalid framing or syntax reserved to GridGrind's parseable formula input. */
public final class InvalidFormulaInputException extends IllegalArgumentException
    implements FormulaInputException {
  private static final long serialVersionUID = 1L;
  private final String publicMessage;
  private final dev.erst.gridgrind.core.problem.PayloadLocation fieldLocation;

  /** Creates one failure with a public, request-safe explanation. */
  public InvalidFormulaInputException(String message) {
    super(Objects.requireNonNull(message, "message must not be null"));
    this.publicMessage = message;
    this.fieldLocation = new dev.erst.gridgrind.core.problem.PayloadLocation.Unavailable();
  }

  /** Creates a classified failure at the caller-owned relative field path. */
  public InvalidFormulaInputException(String message, String fieldPath) {
    super(Objects.requireNonNull(message, "message must not be null"));
    this.publicMessage = message;
    this.fieldLocation =
        new dev.erst.gridgrind.core.problem.PayloadLocation.PathOnly(
            Objects.requireNonNull(fieldPath, "fieldPath must not be null"));
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
