package dev.erst.gridgrind.excel;

import java.nio.file.Path;
import java.util.Objects;

/** Signals that a source file is not an openable OOXML workbook package. */
public final class WorkbookNotOpenableException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  /** Creates one request-safe source-package failure. */
  public WorkbookNotOpenableException(Path workbookPath, Throwable cause) {
    super(
        "Workbook package is not openable: " + Objects.requireNonNull(workbookPath, "workbookPath"),
        cause);
  }
}
