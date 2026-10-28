package dev.erst.gridgrind.engine.formulas;

/** Signals a recognized formula construct rejected by the POI authoring path. */
public final class UnsupportedFormulaConstructException extends IllegalArgumentException
    implements FormulaException {
  private static final long serialVersionUID = 1L;

  private final String sheetName;
  private final String address;
  private final String formula;

  /** Creates the exception for one unsupported authored formula construct. */
  public UnsupportedFormulaConstructException(
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
