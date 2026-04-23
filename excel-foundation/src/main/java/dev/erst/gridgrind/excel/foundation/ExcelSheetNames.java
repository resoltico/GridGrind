package dev.erst.gridgrind.excel.foundation;

import java.util.Objects;

/** Shared Excel sheet-name validation rules reused across protocol and engine surfaces. */
public final class ExcelSheetNames {
  private ExcelSheetNames() {}

  /** Validates one sheet name against GridGrind's Excel-facing contract. */
  public static void requireValid(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    if (value.length() > 31) {
      throw new IllegalArgumentException(fieldName + " must not exceed 31 characters: " + value);
    }
    if (value.charAt(0) == '\'' || value.charAt(value.length() - 1) == '\'') {
      throw new IllegalArgumentException(
          fieldName + " must not begin or end with a single quote: " + value);
    }

    for (int index = 0; index < value.length(); index++) {
      char current = value.charAt(index);
      if (isInvalidExcelSheetCharacter(current)) {
        throw new IllegalArgumentException(
            fieldName
                + " contains invalid Excel character "
                + displayCharacter(current)
                + " at position "
                + (index + 1)
                + ": "
                + value);
      }
    }
  }

  private static boolean isInvalidExcelSheetCharacter(char value) {
    return switch (value) {
      case 0x0000, 0x0003, ':', '\\', '/', '?', '*', '[', ']' -> true;
      default -> false;
    };
  }

  private static String displayCharacter(char value) {
    if (Character.isISOControl(value)) {
      return "U+%04X".formatted((int) value);
    }
    return "'" + value + "'";
  }
}
