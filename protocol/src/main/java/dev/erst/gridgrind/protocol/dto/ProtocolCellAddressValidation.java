package dev.erst.gridgrind.protocol.dto;

import java.util.Locale;
import java.util.Objects;

/** Protocol-owned validation helpers for single-cell A1-style addresses. */
public final class ProtocolCellAddressValidation {
  private static final int MAX_COLUMN_INDEX = 16_384;
  private static final int MAX_ROW_INDEX = 1_048_576;

  private ProtocolCellAddressValidation() {}

  /** Validates one single-cell A1-style address and returns it unchanged. */
  public static String validateAddress(String address) {
    Objects.requireNonNull(address, "address must not be null");
    if (address.isBlank()) {
      throw new IllegalArgumentException("address must not be blank");
    }
    if (!address.matches("(?i)^\\$?[A-Z]{1,3}\\$?[1-9][0-9]*$")) {
      throw new IllegalArgumentException("address must be a single-cell A1-style address");
    }

    String normalized = address.replace("$", "").toUpperCase(Locale.ROOT);
    int splitIndex = 0;
    while (Character.isLetter(normalized.charAt(splitIndex))) {
      splitIndex++;
    }

    String columnLabel = normalized.substring(0, splitIndex);
    int rowNumber = Integer.parseInt(normalized.substring(splitIndex));
    if (columnNumber(columnLabel) > MAX_COLUMN_INDEX || rowNumber > MAX_ROW_INDEX) {
      throw new IllegalArgumentException("address must be within Excel .xlsx bounds");
    }
    return address;
  }

  private static int columnNumber(String columnLabel) {
    int value = 0;
    for (int index = 0; index < columnLabel.length(); index++) {
      value = (value * 26) + (columnLabel.charAt(index) - 'A' + 1);
    }
    return value;
  }
}
