package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** Canonical validation for the two V2 OOXML formula-body input contracts. */
public final class FormulaTextValidation {
  private FormulaTextValidation() {}

  /** Validates the parseable {@code FORMULA} body before it reaches the authoring path. */
  public static String requireNormalFormulaBody(String formula, String fieldName) {
    String normalized = requireNonBlank(formula, fieldName, InvalidFormulaInputException::new);
    if (normalized.startsWith("=")) {
      throw new InvalidFormulaInputException(
          fieldName + " must be an OOXML formula body and must not begin with =");
    }
    if (normalized.startsWith("{=") && normalized.endsWith("}")) {
      throw new InvalidFormulaInputException(
          fieldName + " must be scalar; use SET_ARRAY_FORMULA for an array formula");
    }
    return normalized;
  }

  /** Validates opaque {@code RAW_FORMULA} character data without interpreting formula syntax. */
  public static String requireRawFormulaBody(String formula, String fieldName) {
    String normalized = requireNonBlank(formula, fieldName, InvalidRawFormulaTextException::new);
    if (normalized.startsWith("=")) {
      throw new InvalidRawFormulaTextException(
          fieldName + " must be an OOXML formula body and must not begin with =");
    }
    requireXml10CharacterData(normalized);
    return normalized;
  }

  /** Rejects every Unicode code point that XML 1.0 forbids in character data. */
  public static void requireXml10CharacterData(String formula) {
    Objects.requireNonNull(formula, "formula must not be null");
    for (int index = 0;
        index < formula.length();
        index += Character.charCount(formula.codePointAt(index))) {
      int codePoint = formula.codePointAt(index);
      if (!isXml10Character(codePoint)) {
        throw new InvalidRawFormulaTextException(
            "formula text contains a character not permitted by XML 1.0 at code-point index "
                + index);
      }
    }
  }

  private static boolean isXml10Character(int codePoint) {
    return codePoint == 0x9
        || codePoint == 0xA
        || codePoint == 0xD
        || (codePoint >= 0x20 && codePoint <= 0xD7FF)
        || (codePoint >= 0xE000 && codePoint <= 0xFFFD)
        || codePoint >= 0x10000;
  }

  private static String requireNonBlank(
      String formula,
      String fieldName,
      java.util.function.Function<String, ? extends IllegalArgumentException> failure) {
    Objects.requireNonNull(formula, fieldName + " must not be null");
    if (formula.isBlank()) {
      throw failure.apply(fieldName + " must not be blank");
    }
    return formula;
  }
}
