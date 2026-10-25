package dev.erst.gridgrind.core.vocabulary;

/** One local A1-style cell address within Excel .xlsx bounds. */
public record CellAddress(String value) implements WireString {
  public CellAddress {
    value = ProtocolCellAddressValidation.validateAddress(value);
  }

  @Override
  public String toString() {
    return value;
  }
}
