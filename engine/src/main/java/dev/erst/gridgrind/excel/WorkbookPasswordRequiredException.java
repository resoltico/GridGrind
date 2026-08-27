package dev.erst.gridgrind.excel;

/** Signals that an encrypted OOXML workbook requires a password before it can be opened. */
public final class WorkbookPasswordRequiredException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  /**
   * Creates one semantic password-required failure without exposing private materialization paths.
   */
  public WorkbookPasswordRequiredException() {
    super("The source workbook is encrypted and requires source.security.password before opening");
  }
}
