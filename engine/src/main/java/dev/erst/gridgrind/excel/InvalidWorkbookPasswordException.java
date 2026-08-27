package dev.erst.gridgrind.excel;

/** Signals that the supplied encrypted-workbook password was incorrect. */
public final class InvalidWorkbookPasswordException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  /** Creates one semantic password failure without exposing private materialization paths. */
  public InvalidWorkbookPasswordException() {
    super("The supplied source.security.password did not unlock the source workbook");
  }
}
