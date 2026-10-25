package dev.erst.gridgrind.contract.dto;

/** Shared dangerous-formula rejection applied to all formula-bearing input types. */
final class FormulaInputSecurity {
  private FormulaInputSecurity() {}

  static void rejectDde(String formula) { // LIM-027, LIM-031, LIM-032, LIM-033, LIM-034
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
    if (toCheck.regionMatches(true, 0, "webservice(", 0, 11)) { // LIM-031
      throw new IllegalArgumentException(
          "formula must not use the WEBSERVICE function;"
              + " WEBSERVICE triggers outbound HTTP requests when the workbook is opened in Excel");
    }
  }
}
