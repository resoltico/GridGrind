package dev.erst.gridgrind.excel;

/** Signals that a requested cell address does not currently exist in the sheet. */
public final class CellNotFoundException extends java.util.NoSuchElementException {
  private static final long serialVersionUID = 1L;

  private final String address;

  /** Creates the exception for the given A1-style cell address. */
  public CellNotFoundException(String address) {
    super("Cell does not exist: " + address);
    this.address = address;
  }

  public String address() {
    return address;
  }
}
