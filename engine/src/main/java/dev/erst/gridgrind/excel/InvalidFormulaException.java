package dev.erst.gridgrind.excel;

/** Signals that a formula cannot be parsed or resolved by the current workbook context. */
public final class InvalidFormulaException extends IllegalArgumentException
    implements FormulaException {
  private static final long serialVersionUID = 1L;

  private final String sheetName;
  private final String address;
  private final String formula;

  /** Creates the exception for the formula at the given sheet location. */
  public InvalidFormulaException(
      String sheetName, String address, String formula, String message, Throwable cause) {
    super(message, cause);
    this.sheetName = sheetName;
    this.address = address;
    this.formula = formula;
  }

  @Override
  public String sheetName() {
    return sheetName;
  }

  @Override
  public String address() {
    return address;
  }

  @Override
  public String formula() {
    return formula;
  }
}
