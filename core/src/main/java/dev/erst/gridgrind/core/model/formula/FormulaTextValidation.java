package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.source.OoxmlTextValidation;
import java.util.Objects;

/** Canonical validation for normal and raw OOXML formula-body authoring contracts. */
public final class FormulaTextValidation {
  private FormulaTextValidation() {}

  /** Validates the parseable {@code FORMULA} body before it reaches the authoring path. */
  public static String requireNormalFormulaBody(String formula, String fieldName) {
    String normalized =
        requireNonBlank(
            formula, fieldName, message -> new InvalidFormulaInputException(message, fieldName));
    if (normalized.startsWith("=")) {
      throw new InvalidFormulaInputException(
          fieldName + " must be an OOXML formula body and must not begin with =", fieldName);
    }
    if (normalized.startsWith("{=") && normalized.endsWith("}")) {
      throw new InvalidFormulaInputException(
          fieldName + " must be scalar; use SET_ARRAY_FORMULA for an array formula", fieldName);
    }
    return normalized;
  }

  /** Validates opaque {@code RAW_FORMULA} character data without interpreting formula syntax. */
  public static String requireRawFormulaBody(String formula, String fieldName) {
    String normalized =
        requireNonBlank(
            formula, fieldName, message -> new InvalidRawFormulaTextException(message, fieldName));
    if (normalized.startsWith("=")) {
      throw new InvalidRawFormulaTextException(
          fieldName + " must be an OOXML formula body and must not begin with =", fieldName);
    }
    try {
      requireXml10CharacterData(normalized);
    } catch (InvalidRawFormulaTextException invalid) {
      throw new InvalidRawFormulaTextException(invalid.publicMessage(), fieldName, invalid);
    }
    return normalized;
  }

  /** Rejects every Unicode code point that XML 1.0 forbids in character data. */
  public static void requireXml10CharacterData(String formula) {
    try {
      OoxmlTextValidation.requireXml10CharacterData(formula, "formula text");
    } catch (IllegalArgumentException exception) {
      throw new InvalidRawFormulaTextException(
          Objects.requireNonNull(exception.getMessage(), "validation message must not be null"),
          exception);
    }
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
