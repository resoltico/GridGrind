package dev.erst.gridgrind.excel;

import java.io.IOException;

/** Signals an unexpected OOXML package-security runtime failure. */
public final class WorkbookSecurityException extends IOException {
  private static final long serialVersionUID = 1L;

  /** Creates the exception for the given security-runtime failure. */
  public WorkbookSecurityException(String message) {
    super(message);
  }

  /** Creates the exception for the given security-runtime failure. */
  public WorkbookSecurityException(String message, Throwable cause) {
    super(message, cause);
  }
}
