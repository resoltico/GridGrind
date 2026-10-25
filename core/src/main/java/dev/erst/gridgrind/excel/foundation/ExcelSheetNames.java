package dev.erst.gridgrind.excel.foundation;

import java.util.Objects;
import java.util.Optional;

/** Shared Excel sheet-name validation rules reused across protocol and engine surfaces. */
public final class ExcelSheetNames {
  private static final String INVALID_EXCEL_SHEET_CHARACTERS = ":\\/?*[]";

  private ExcelSheetNames() {}

  /** Validates one sheet name against GridGrind's Excel-facing contract. */
  public static void requireValid(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    violation(value)
        .ifPresent(
            violation -> {
              throw new IllegalArgumentException(violation.message(fieldName));
            });
  }

  /** Returns the first sheet-name rule violation, if any. */
  public static Optional<ExcelSheetNameProblem> violation(String value) {
    Objects.requireNonNull(value, "value must not be null");
    if (value.isBlank()) {
      return Optional.of(ExcelSheetNameProblem.blank());
    }
    if (value.length() > 31) {
      return Optional.of(ExcelSheetNameProblem.tooLong(value));
    }
    if (value.charAt(0) == '\'' || value.charAt(value.length() - 1) == '\'') {
      return Optional.of(ExcelSheetNameProblem.boundaryQuote(value));
    }
    for (int index = 0; index < value.length(); index++) {
      char current = value.charAt(index);
      if (isInvalidExcelSheetCharacter(current)) {
        return Optional.of(
            ExcelSheetNameProblem.invalidCharacter(displayCharacter(current), index + 1, value));
      }
    }
    return Optional.empty();
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
}
