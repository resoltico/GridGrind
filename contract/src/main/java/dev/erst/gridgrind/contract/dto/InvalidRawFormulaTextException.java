package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Signals invalid OOXML formula character data supplied through the opaque formula input. */
public final class InvalidRawFormulaTextException extends IllegalArgumentException
    implements FormulaInputException {
  private static final long serialVersionUID = 1L;
  private final String publicMessage;

  /** Creates one failure with a public, request-safe explanation. */
  public InvalidRawFormulaTextException(String message) {
    super(Objects.requireNonNull(message, "message must not be null"));
    this.publicMessage = message;
  }

  @Override
  public String publicMessage() {
    return publicMessage;
  }
}
