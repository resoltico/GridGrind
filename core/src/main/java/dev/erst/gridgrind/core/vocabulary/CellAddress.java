package dev.erst.gridgrind.core.vocabulary;

import dev.erst.gridgrind.core.limits.Limit;

/** One concrete local cell position, with zero-based indices and canonical A1 presentation. */
public record CellAddress(int rowIndex, int columnIndex) implements WireString {
  public CellAddress {
    if (rowIndex < 0
        || rowIndex > Limit.MAX_ROW_INDEX
        || columnIndex < 0
        || columnIndex > Limit.MAX_COLUMN_INDEX) {
      throw ProtocolCellAddressValidation.invalid(ProtocolCellAddressValidation.Violation.BOUNDS);
    }
  }

  /** Parses one local A1 address; case and absolute markers do not affect position identity. */
  public CellAddress(String address) {
    this(ProtocolCellAddressValidation.parseAddress(address));
  }

  private CellAddress(CellAddress address) {
    this(address.rowIndex(), address.columnIndex());
  }

  @Override
  public String value() {
    int column = columnIndex + 1;
    StringBuilder label = new StringBuilder();
    while (column > 0) {
      label.append((char) ('A' + (column - 1) % 26));
      column = (column - 1) / 26;
    }
    return label.reverse().append(rowIndex + 1).toString();
  }

  @Override
  public String toString() {
    return value();
  }
}
