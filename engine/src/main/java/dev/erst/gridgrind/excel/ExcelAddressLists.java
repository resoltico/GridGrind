package dev.erst.gridgrind.excel;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Set;

/** Canonical validation for ordered address lists used by exact-cell read surfaces. */
public final class ExcelAddressLists {
  private ExcelAddressLists() {}

  /** Copies one non-empty ordered address list after rejecting blanks, nulls, and duplicates. */
  public static List<String> copyNonEmptyDistinctAddresses(List<String> addresses) { // LIM-007
    Objects.requireNonNull(addresses, "addresses must not be null");
    List<String> copy = List.copyOf(addresses);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("addresses must not be empty");
    }
    Set<String> unique = new LinkedHashSet<>();
    for (int index = 0; index < copy.size(); index++) {
      String validated;
      try {
        validated = validateAddress(copy.get(index));
      } catch (IllegalArgumentException exception) {
        throw new IllegalArgumentException(
            "addresses[" + index + "] " + exception.getMessage(), exception);
      }
      if (!unique.add(validated.toUpperCase(Locale.ROOT))) {
        throw new IllegalArgumentException("addresses must not contain duplicates");
      }
    }
    return copy;
  }

  private static String validateAddress(String address) {
    Objects.requireNonNull(address, "addresses must not contain nulls");
    if (address.isBlank()) {
      throw new IllegalArgumentException("addresses must not contain blank values");
    }
    if (!address.matches("(?i)^\\$?[A-Z]{1,3}\\$?[1-9][0-9]*$")) {
      throw new IllegalArgumentException("address must be a single-cell A1-style address");
    }
    String normalized = address.replace("$", "").toUpperCase(Locale.ROOT);
    int splitIndex = 0;
    while (Character.isLetter(normalized.charAt(splitIndex))) {
      splitIndex++;
    }
    if (columnNumber(normalized.substring(0, splitIndex)) > 16_384
        || Integer.parseInt(normalized.substring(splitIndex)) > 1_048_576) {
      throw new IllegalArgumentException("addresses must stay within Excel .xlsx bounds");
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
