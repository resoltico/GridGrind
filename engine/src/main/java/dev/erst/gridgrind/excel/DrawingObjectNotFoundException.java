package dev.erst.gridgrind.excel;

/** Thrown when a requested drawing object cannot be located on a sheet. */
public final class DrawingObjectNotFoundException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  /** Creates the exception for one missing sheet-local drawing object name. */
  public DrawingObjectNotFoundException(String sheetName, String objectName) {
    super("Drawing object '" + objectName + "' was not found on sheet '" + sheetName + "'");
  }
}
