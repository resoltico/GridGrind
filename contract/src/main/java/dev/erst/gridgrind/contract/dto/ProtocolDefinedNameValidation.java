package dev.erst.gridgrind.contract.dto;

import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

/** Protocol-owned validation helpers for defined names and related identifiers. */
public final class ProtocolDefinedNameValidation {
  private ProtocolDefinedNameValidation() {}

  /** Validates one protocol-facing defined-name identifier and returns its canonical text. */
  public static String validateName(String name) {
    Objects.requireNonNull(name, "name must not be null");
    violation(name)
        .ifPresent(
            violation -> {
              throw new IllegalArgumentException(renderViolation(violation));
            });
    return name;
  }

  /** Returns the first defined-name rule violation, if any. */
  public static Optional<Violation> violation(String name) {
    Objects.requireNonNull(name, "name must not be null");
    if (name.isBlank()) {
      return Optional.of(Violation.BLANK);
    }
    if (name.codePointCount(0, name.length())
        > ProtocolConstraintValues.DEFINED_NAME_MAX_CODE_POINTS) {
      return Optional.of(Violation.TOO_LONG);
    }
    if (!isSupportedIdentifier(name)) {
      return Optional.of(Violation.SYNTAX);
    }
    if (name.regionMatches(true, 0, "_xlnm.", 0, "_xlnm.".length())) {
      return Optional.of(Violation.RESERVED_PREFIX);
    }
    if (looksLikeA1CellReference(name)) {
      return Optional.of(Violation.A1_COLLISION);
    }
    if (name.matches("(?i)^R[1-9][0-9]*C[1-9][0-9]*$")) {
      return Optional.of(Violation.R1C1_COLLISION);
    }
    return Optional.empty();
  }

  private static boolean looksLikeA1CellReference(String candidate) {
    if (!candidate.matches("(?i)^\\$?[A-Z]{1,3}\\$?[1-9][0-9]*$")) {
      return false;
    }
    String normalized = candidate.replace("$", "").toUpperCase(Locale.ROOT);
    String columnLabel = normalized.replaceAll("\\d.*$", "");
    int columnNumber = 0;
    for (int index = 0; index < columnLabel.length(); index++) {
      columnNumber = (columnNumber * 26) + (columnLabel.charAt(index) - 'A' + 1);
    }
    return columnNumber <= 16384;
  }

  private static boolean isSupportedIdentifier(String name) {
    int firstCodePoint = name.codePointAt(0);
    if (!Character.isLetter(firstCodePoint) && firstCodePoint != '_' && firstCodePoint != '\\') {
      return false;
    }
    int offset = Character.charCount(firstCodePoint);
    while (offset < name.length()) {
      int codePoint = name.codePointAt(offset);
      if (!Character.isLetter(codePoint)
          && !isUnicodeNumber(codePoint)
          && codePoint != '_'
          && codePoint != '.'
          && codePoint != '\\') {
        return false;
      }
      offset += Character.charCount(codePoint);
    }
    return true;
  }

  private static boolean isUnicodeNumber(int codePoint) {
    return switch (Character.getType(codePoint)) {
      case Character.DECIMAL_DIGIT_NUMBER, Character.LETTER_NUMBER, Character.OTHER_NUMBER -> true;
      default -> false;
    };
  }

  private static String renderViolation(Violation violation) {
    return switch (violation) {
      case BLANK -> "name must not be blank";
      case TOO_LONG -> "name must not exceed 255 Unicode code points";
      case SYNTAX ->
          "name must start with a letter, underscore, or backslash and contain only Unicode"
              + " letters, Unicode numbers, underscore, period, or backslash";
      case RESERVED_PREFIX -> "name must not use the reserved _xlnm. prefix";
      case A1_COLLISION -> "name must not collide with A1-style cell reference syntax";
      case R1C1_COLLISION -> "name must not collide with R1C1-style cell reference syntax";
    };
  }

  /** Structured defined-name violation family. */
  public enum Violation {
    BLANK,
    TOO_LONG,
    SYNTAX,
    RESERVED_PREFIX,
    A1_COLLISION,
    R1C1_COLLISION
  }
}
