package dev.erst.gridgrind.engine.runtime;

/** Raised when an authored workbook source path names a directory rather than a workbook file. */
final class SourcePathIsDirectoryException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  SourcePathIsDirectoryException(String path) {
    super("Workbook source path is a directory: " + path);
  }
}
