package dev.erst.gridgrind.excel;

/** Signals that a rectangular A1-style range address is invalid. */
public final class InvalidRangeAddressException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  private final String range;

  /** Creates the exception for the given range address and the original parsing cause. */
  public InvalidRangeAddressException(String range, Throwable cause) {
    super("Invalid range address: " + range, cause);
    this.range = range;
  }

  public String range() {
    return range;
  }
}
