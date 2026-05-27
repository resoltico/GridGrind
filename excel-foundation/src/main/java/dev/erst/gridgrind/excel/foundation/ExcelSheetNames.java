package dev.erst.gridgrind.excel.foundation;

import java.util.Objects;

/** Shared Excel sheet-name validation rules reused across protocol and engine surfaces. */
public final class ExcelSheetNames {
  private static final String INVALID_EXCEL_SHEET_CHARACTERS = ":\\/?*[]";

  private ExcelSheetNames() {}

  /** Validates one sheet name against GridGrind's Excel-facing contract. */
  public static void requireValid(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    requireNonBlank(value, fieldName);
    requireLengthWithinExcelLimit(value, fieldName);
    requireNoBoundaryQuote(value, fieldName);
    requireOnlyValidExcelCharacters(value, fieldName);
  }

  private static boolean isInvalidExcelSheetCharacter(char value) {
    return value == 0x0000 || value == 0x0003 || INVALID_EXCEL_SHEET_CHARACTERS.indexOf(value) >= 0;
  }

  private static String displayCharacter(char value) {
    if (Character.isISOControl(value)) {
      return "U+%04X".formatted((int) value);
    }
    return "'" + value + "'";
  }

  private static void requireNonBlank(String value, String fieldName) {
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
  }

  private static void requireLengthWithinExcelLimit(String value, String fieldName) {
    if (value.length() > 31) {
      throw new IllegalArgumentException(fieldName + " must not exceed 31 characters: " + value);
    }
  }

  private static void requireNoBoundaryQuote(String value, String fieldName) {
    if (value.charAt(0) == '\'' || value.charAt(value.length() - 1) == '\'') {
      throw new IllegalArgumentException(
          fieldName + " must not begin or end with a single quote: " + value);
    }
  }

  private static void requireOnlyValidExcelCharacters(String value, String fieldName) {
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
}
