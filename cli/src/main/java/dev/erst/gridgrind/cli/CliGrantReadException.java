package dev.erst.gridgrind.cli;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Objects;

/** Signals that one command-side host grant document could not be read or decoded. */
final class CliGrantReadException extends RuntimeException {
  private static final long serialVersionUID = 1L;
  private final Path grantPath;

  CliGrantReadException(Path grantPath, IOException cause) {
    super(Objects.requireNonNull(cause, "cause must not be null").getMessage(), cause);
    this.grantPath = Objects.requireNonNull(grantPath, "grantPath must not be null");
  }

  Path grantPath() {
    return grantPath;
  }
}
