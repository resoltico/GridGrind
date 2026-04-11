package dev.erst.gridgrind.excel;

/** Signals that formula evaluation required a user-defined function that was not registered. */
public final class UnregisteredUserDefinedFunctionException extends IllegalArgumentException
    implements FormulaException {
  private static final long serialVersionUID = 1L;

  private final String sheetName;
  private final String address;
  private final String formula;
  private final String functionName;

  /** Creates the exception for one missing user-defined function registration. */
  public UnregisteredUserDefinedFunctionException(
      String sheetName,
      String address,
      String formula,
      String functionName,
      String message,
      Throwable cause) {
    super(message, cause);
    this.sheetName = sheetName;
    this.address = address;
    this.formula = formula;
    this.functionName = functionName;
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

  /** Returns the missing user-defined function name. */
  public String functionName() {
    return functionName;
  }
}
