package dev.erst.gridgrind.excel;

/** Signals that a workbook does not contain the requested sheet. */
public final class SheetNotFoundException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  private final String sheetName;

  /** Creates the exception for the given sheet name. */
  public SheetNotFoundException(String sheetName) {
    super("Sheet does not exist: " + sheetName);
    this.sheetName = sheetName;
  }

  public String sheetName() {
    return sheetName;
  }
}
