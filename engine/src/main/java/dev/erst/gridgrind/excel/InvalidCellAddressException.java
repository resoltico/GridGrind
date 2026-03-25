package dev.erst.gridgrind.excel;

/** Signals that a cell address is not valid A1 notation for Excel. */
public final class InvalidCellAddressException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  private final String address;

  /** Creates the exception for the given cell address and the original parsing cause. */
  public InvalidCellAddressException(String address, Throwable cause) {
    super("Invalid cell address: " + address, cause);
    this.address = address;
  }

  public String address() {
    return address;
  }
}
