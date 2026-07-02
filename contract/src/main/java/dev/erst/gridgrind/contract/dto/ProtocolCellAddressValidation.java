package dev.erst.gridgrind.contract.dto;

import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

/** Protocol-owned validation helpers for single-cell A1-style addresses. */
public final class ProtocolCellAddressValidation {
  private static final int MAX_COLUMN_INDEX = 16_384;
  private static final int MAX_ROW_INDEX = 1_048_576;

  private ProtocolCellAddressValidation() {}

  /** Validates one single-cell A1-style address and returns it unchanged. */
  public static String validateAddress(String address) {
    Objects.requireNonNull(address, "address must not be null");
    violation(address)
        .ifPresent(
            violation -> {
              throw new IllegalArgumentException(renderViolation(violation));
            });
    return address;
  }

  /** Returns the first address rule violation, if any. */
  public static Optional<Violation> violation(String address) {
    Objects.requireNonNull(address, "address must not be null");
    if (address.isBlank()) {
      return Optional.of(Violation.BLANK);
    }
    if (!address.matches("(?i)^\\$?[A-Z]{1,3}\\$?[1-9][0-9]*$")) {
      return Optional.of(Violation.SYNTAX);
    }

    String normalized = address.replace("$", "").toUpperCase(Locale.ROOT);
    int splitIndex = 0;
    while (Character.isLetter(normalized.charAt(splitIndex))) {
      splitIndex++;
    }

    String columnLabel = normalized.substring(0, splitIndex);
    int rowNumber = Integer.parseInt(normalized.substring(splitIndex));
    if (columnNumber(columnLabel) > MAX_COLUMN_INDEX || rowNumber > MAX_ROW_INDEX) {
      return Optional.of(Violation.BOUNDS);
    }
    return Optional.empty();
  }

  private static String renderViolation(Violation violation) {
    return switch (violation) {
      case BLANK -> "address must not be blank";
      case SYNTAX -> "address must be a single-cell A1-style address";
      case BOUNDS -> "address must be within Excel .xlsx bounds";
    };
  }

  private static int columnNumber(String columnLabel) {
    int value = 0;
    for (int index = 0; index < columnLabel.length(); index++) {
      value = (value * 26) + (columnLabel.charAt(index) - 'A' + 1);
    }
    return value;
  }

  /** Structured address violation family. */
  public enum Violation {
    BLANK,
    SYNTAX,
    BOUNDS
  }
}
