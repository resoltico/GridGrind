package dev.erst.gridgrind.excel;

/** Signals that a formula cannot be parsed or resolved by the current workbook context. */
public final class InvalidFormulaException extends FormulaException {
  private static final long serialVersionUID = 1L;

  /** Creates the exception for the formula at the given sheet location. */
  public InvalidFormulaException(
      String sheetName, String address, String formula, String message, Throwable cause) {
    super(message, sheetName, address, formula, cause);
  }
}
