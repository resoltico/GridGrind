package dev.erst.gridgrind.engine.runtime;

/** Raised when an authored workbook output path names a directory rather than a file leaf. */
final class OutputPathIsDirectoryException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  OutputPathIsDirectoryException(String path) {
    super("Workbook output path is a directory: " + path);
  }
}
