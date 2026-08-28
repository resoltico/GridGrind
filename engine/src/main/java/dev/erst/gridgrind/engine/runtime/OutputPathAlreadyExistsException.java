package dev.erst.gridgrind.engine.runtime;

/** Raised when create-new workbook persistence conflicts with an existing authored output path. */
final class OutputPathAlreadyExistsException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  OutputPathAlreadyExistsException(String path) {
    super("Workbook output path already exists and ifExists=REJECT: " + path);
  }

  OutputPathAlreadyExistsException(String path, Throwable cause) {
    super("Workbook output path already exists and ifExists=REJECT: " + path, cause);
  }
}
