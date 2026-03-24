package dev.erst.gridgrind.excel;

/** Base class for formula problems detected while writing or evaluating workbook formulas. */
public abstract sealed class FormulaException extends IllegalArgumentException
    permits InvalidFormulaException, UnsupportedFormulaException {
  private static final long serialVersionUID = 1L;

  private final String sheetName;
  private final String address;
  private final String formula;

  protected FormulaException(
      String message, String sheetName, String address, String formula, Throwable cause) {
    super(message, cause);
    this.sheetName = sheetName;
    this.address = address;
    this.formula = formula;
  }

  public String sheetName() {
    return sheetName;
  }

  public String address() {
    return address;
  }

  public String formula() {
    return formula;
  }
}
