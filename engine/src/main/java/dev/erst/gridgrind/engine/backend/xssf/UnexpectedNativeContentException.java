package dev.erst.gridgrind.engine.backend.xssf;

import java.util.Objects;

/** Keeps an unclassified native-reader failure unexpected, including native argument exceptions. */
public final class UnexpectedNativeContentException extends RuntimeException {
  private static final long serialVersionUID = 1L;

  /** Preserves the original implementation failure without treating it as caller input. */
  public UnexpectedNativeContentException(Exception cause) {
    super(
        "Unexpected native content failure",
        Objects.requireNonNull(cause, "cause must not be null"));
  }
}
