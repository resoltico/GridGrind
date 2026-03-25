package dev.erst.gridgrind.excel;

import java.nio.file.Path;

/** Signals that a workbook file path does not exist on disk. */
public final class WorkbookNotFoundException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  private final Path workbookPath;

  /** Creates the exception for the given workbook path. */
  public WorkbookNotFoundException(Path workbookPath) {
    super("Workbook does not exist: " + workbookPath);
    this.workbookPath = workbookPath;
  }

  public Path workbookPath() {
    return workbookPath;
  }
}
