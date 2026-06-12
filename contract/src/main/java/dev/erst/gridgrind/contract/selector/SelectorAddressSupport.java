package dev.erst.gridgrind.contract.selector;

import java.util.Locale;

/** Shared A1-style address math for selector families. */
final class SelectorAddressSupport {
  private SelectorAddressSupport() {}

  static String absoluteA1Address(int rowIndex, int columnIndex) {
    return columnLabel(columnIndex) + (rowIndex + 1);
  }

  static int columnIndex(String address) {
    String normalized = address.replace("$", "").toUpperCase(Locale.ROOT);
    int splitIndex = 0;
    while (splitIndex < normalized.length() && Character.isLetter(normalized.charAt(splitIndex))) {
      splitIndex++;
    }
    int result = 0;
    for (int index = 0; index < splitIndex; index++) {
      result = (result * 26) + (normalized.charAt(index) - 'A' + 1);
    }
    return result - 1;
  }

  static int rowIndex(String address) {
    String normalized = address.replace("$", "");
    int splitIndex = 0;
    while (splitIndex < normalized.length() && !Character.isDigit(normalized.charAt(splitIndex))) {
      splitIndex++;
    }
    return Integer.parseInt(normalized.substring(splitIndex)) - 1;
  }

  private static String columnLabel(int columnIndex) {
    int value = columnIndex + 1;
    StringBuilder builder = new StringBuilder();
    while (value > 0) {
      int remainder = (value - 1) % 26;
      builder.append((char) ('A' + remainder));
      value = (value - 1) / 26;
    }
    return builder.reverse().toString();
  }
}
