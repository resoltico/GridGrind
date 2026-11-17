package dev.erst.gridgrind.cli.io;

import java.nio.file.Path;
import java.util.Objects;

/** Retains the authority document role for an unexpected reader or provider failure. */
public final class CliAuthorityUnexpectedException extends RuntimeException {
  private static final long serialVersionUID = 1L;
  private final Path authorityPath;

  CliAuthorityUnexpectedException(Path authorityPath, RuntimeException cause) {
    super("Unexpected authority reader failure", Objects.requireNonNull(cause));
    this.authorityPath = Objects.requireNonNull(authorityPath);
  }

  /** Original authority document path, independent of the request document. */
  public Path authorityPath() {
    return authorityPath;
  }
}
