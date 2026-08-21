package dev.erst.gridgrind.cli;

import java.util.Objects;

/** Signals that the primary output stream may already contain an incomplete payload. */
final class CliPrimaryOutputException extends RuntimeException {
  private static final long serialVersionUID = 1L;

  CliPrimaryOutputException(java.io.IOException cause) {
    super(
        "Primary output transport failed", Objects.requireNonNull(cause, "cause must not be null"));
  }
}
