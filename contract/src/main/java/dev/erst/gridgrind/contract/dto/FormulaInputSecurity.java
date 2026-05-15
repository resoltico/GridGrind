package dev.erst.gridgrind.contract.dto;

/** Shared DDE rejection applied to all formula-bearing input types. */
final class FormulaInputSecurity {
  private FormulaInputSecurity() {}

  static void rejectDde(String formula) { // LIM-027
    String toCheck = formula;
    if (toCheck.startsWith("{=") && toCheck.endsWith("}")) {
      toCheck = toCheck.substring(2, toCheck.length() - 1);
    } else if (toCheck.startsWith("=")) {
      toCheck = toCheck.substring(1);
    }
    if (toCheck.regionMatches(true, 0, "dde(", 0, 4)) { // LIM-027
      throw new IllegalArgumentException(
          "formula must not use the DDE function; DDE formula execution is a security risk");
    }
  }
}
