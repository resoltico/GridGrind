package dev.erst.gridgrind.contract.json;

import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Signals request bytes that cannot be decoded as UTF-8 before JSON parsing begins. */
public final class InvalidEncodingException extends IllegalArgumentException
    implements PayloadException {
  private static final long serialVersionUID = 1L;

  private final PayloadLocation jsonLocation;

  /** Creates the exception with the given message, JSON location, and decoding cause. */
  public InvalidEncodingException(
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
