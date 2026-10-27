package dev.erst.gridgrind.cli.io;

import java.util.Objects;

/** Signals that the primary output stream may already contain an incomplete payload. */
public final class CliPrimaryOutputException extends RuntimeException {
  private static final long serialVersionUID = 1L;

  /** Retains a primary-output failure so the dispatcher avoids a second write. */
  public CliPrimaryOutputException(java.io.IOException cause) {
    super(
        "Primary output transport failed", Objects.requireNonNull(cause, "cause must not be null"));
  }
}
