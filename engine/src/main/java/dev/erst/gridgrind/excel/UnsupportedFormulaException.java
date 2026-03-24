package dev.erst.gridgrind.excel;

/** Signals that a formula uses functionality unsupported by the current evaluator. */
public final class UnsupportedFormulaException extends FormulaException {
  private static final long serialVersionUID = 1L;

  public UnsupportedFormulaException(
      String sheetName, String address, String formula, String message, Throwable cause) {
    super(message, sheetName, address, formula, cause);
  }
}
