package dev.erst.gridgrind.excel;

import java.nio.file.Path;

/** Signals that an encrypted OOXML workbook requires a password before it can be opened. */
public final class WorkbookPasswordRequiredException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  private final Path workbookPath;

  /** Creates the exception for the encrypted workbook path. */
  public WorkbookPasswordRequiredException(Path workbookPath) {
    super(
        "The workbook is encrypted and requires source.security.password before it can be opened: "
            + workbookPath);
    this.workbookPath = workbookPath;
  }

  public Path workbookPath() {
    return workbookPath;
  }
}
