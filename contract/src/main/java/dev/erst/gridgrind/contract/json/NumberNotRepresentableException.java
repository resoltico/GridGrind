package dev.erst.gridgrind.contract.json;

import java.util.Optional;

/** Signals that one JSON numeric token would lose information in an IEEE-backed request field. */
public final class NumberNotRepresentableException extends IllegalArgumentException
    implements PayloadException {
  private static final long serialVersionUID = 1L;

  private final PayloadLocation jsonLocation;

  /** Creates one exactness failure at the authored request path. */
  public NumberNotRepresentableException(String message, String jsonPath) {
    super(message);
    this.jsonLocation =
        PayloadLocation.from(Optional.of(jsonPath), Optional.empty(), Optional.empty());
  }

  @Override
  public PayloadLocation jsonLocation() {
    return jsonLocation;
  }
}
