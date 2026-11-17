package dev.erst.gridgrind.cli.io;

import dev.erst.gridgrind.core.problem.InvalidAuthorityException;
import java.nio.file.Path;
import java.util.Objects;

/** Adds the actual host authority-file reference to an already classified document refusal. */
public final class CliAuthorityDocumentException extends RuntimeException {
  private static final long serialVersionUID = 1L;
  private final Path authorityPath;
  private final InvalidAuthorityException refusal;

  /** Retains document provenance without reclassifying parser or model faults. */
  public CliAuthorityDocumentException(Path path, InvalidAuthorityException refusal) {
    super(Objects.requireNonNull(refusal).getMessage(), refusal);
    authorityPath = Objects.requireNonNull(path);
    this.refusal = refusal;
  }

  /** Returns the actual selected grant file. */
  public Path authorityPath() {
    return authorityPath;
  }

  /** Returns the fail-closed strict document refusal. */
  public InvalidAuthorityException refusal() {
    return refusal;
  }
}
