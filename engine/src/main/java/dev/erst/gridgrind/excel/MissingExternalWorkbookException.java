package dev.erst.gridgrind.excel;

/** Signals that formula evaluation required an external workbook that was missing or unbound. */
public final class MissingExternalWorkbookException extends IllegalArgumentException
    implements FormulaException {
  private static final long serialVersionUID = 1L;

  private final String sheetName;
  private final String address;
  private final String formula;
  private final String workbookName;

  /** Creates the exception for one missing external workbook reference. */
  public MissingExternalWorkbookException(
      String sheetName,
      String address,
      String formula,
      String workbookName,
      String message,
      Throwable cause) {
    super(message, cause);
    this.sheetName = sheetName;
    this.address = address;
    this.formula = formula;
    this.workbookName = workbookName;
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

  /** Returns the external workbook name referenced by the formula. */
  public String workbookName() {
    return workbookName;
  }
}
