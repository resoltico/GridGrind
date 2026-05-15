package dev.erst.gridgrind.contract.json;

import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Signals that valid JSON did not match the GridGrind request or response payload shape. */
public final class InvalidRequestShapeException extends IllegalArgumentException
    implements PayloadException {
  private static final long serialVersionUID = 1L;

  private final PayloadLocation jsonLocation;

  /** Creates the exception with the given message, JSON location, and binding cause. */
  public InvalidRequestShapeException(
      String message,
      Optional<String> jsonPath,
      Optional<Integer> jsonLine,
      Optional<Integer> jsonColumn,
      @Nullable Throwable cause) {
    super(message, cause);
    this.jsonLocation = PayloadLocation.from(jsonPath, jsonLine, jsonColumn);
  }

  @Override
  public PayloadLocation jsonLocation() {
    return jsonLocation;
  }
}
