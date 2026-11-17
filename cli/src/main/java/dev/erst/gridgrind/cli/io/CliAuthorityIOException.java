package dev.erst.gridgrind.cli.io;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Objects;

/** Authority-file access is an environment/I/O failure, not an invalid authority document. */
public final class CliAuthorityIOException extends IOException {
  private static final long serialVersionUID = 1L;
  private final Path authorityPath;

  /** Retains the selected host file and original private I/O cause. */
  public CliAuthorityIOException(Path path, IOException cause) {
    super("Could not read the selected authority document.", Objects.requireNonNull(cause));
    authorityPath = Objects.requireNonNull(path);
  }

  /** Returns the actual selected grant file. */
  public Path authorityPath() {
    return authorityPath;
  }
}
