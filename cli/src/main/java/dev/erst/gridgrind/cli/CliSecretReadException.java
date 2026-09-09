package dev.erst.gridgrind.cli;

import java.io.IOException;
import java.util.Objects;

/** Signals a host-owned secret-provider read failure without carrying secret material. */
final class CliSecretReadException extends RuntimeException {
  private static final long serialVersionUID = 1L;
  private final String referenceId;

  CliSecretReadException(String referenceId, IOException cause) {
    super(
        "Unable to resolve secret reference "
            + Objects.requireNonNull(referenceId, "referenceId must not be null"),
        cause);
    this.referenceId = referenceId;
  }

  String referenceId() {
    return referenceId;
  }
}
