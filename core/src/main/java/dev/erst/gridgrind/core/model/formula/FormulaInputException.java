package dev.erst.gridgrind.core.model.formula;

/** Shared contract for classified formula-input validation failures. */
public sealed interface FormulaInputException
    permits InvalidFormulaInputException, InvalidRawFormulaTextException {
  /** Returns the non-null request-safe message for the input failure. */
  String publicMessage();

  /** Relative owned field path when the validating caller supplies one. */
  java.util.Optional<String> fieldPath();
}
