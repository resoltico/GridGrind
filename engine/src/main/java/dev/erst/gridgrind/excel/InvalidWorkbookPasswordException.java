package dev.erst.gridgrind.excel;

import java.nio.file.Path;

/** Signals that the supplied encrypted-workbook password was incorrect. */
public final class InvalidWorkbookPasswordException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  private final Path workbookPath;

  /** Creates the exception for the encrypted workbook path. */
  public InvalidWorkbookPasswordException(Path workbookPath) {
    super("The supplied source.security.password did not unlock the workbook: " + workbookPath);
    this.workbookPath = workbookPath;
  }

  public Path workbookPath() {
    return workbookPath;
  }
}
