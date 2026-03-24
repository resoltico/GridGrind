package dev.erst.gridgrind.excel;

/**
 * Translates raw Apache POI formula evaluation exceptions into typed GridGrind exception hierarchy
 * entries.
 */
final class FormulaExceptions {
  private FormulaExceptions() {}

  /**
   * Wraps a raw POI runtime exception thrown during formula evaluation or parsing into a typed
   * {@link InvalidFormulaException} or {@link UnsupportedFormulaException} when the exception
   * originates from a known POI formula package. Returns the original exception unchanged for all
   * other exception types.
   *
   * @param sheetName the name of the sheet containing the formula cell
   * @param address the A1-style address of the formula cell
   * @param formula the raw formula expression that triggered the exception
   * @param exception the raw exception thrown by POI
   * @return a typed formula exception, or the original exception if it is not a formula error
   */
  static RuntimeException wrap(
      String sheetName, String address, String formula, RuntimeException exception) {
    String className = exception.getClass().getName();
    if (className.contains(".ss.formula.eval.")) {
      if (className.contains("NotImplemented")) {
        return new UnsupportedFormulaException(
            sheetName,
            address,
            formula,
            "Unsupported formula"
                + functionLabel(formula)
                + " at "
                + location(sheetName, address)
                + ": "
                + formula,
            exception);
      }
      return exception;
    }
    if (className.startsWith("org.apache.poi.ss.formula.")) {
      return new InvalidFormulaException(
          sheetName,
          address,
          formula,
          "Invalid formula at " + location(sheetName, address) + ": " + formula,
          exception);
    }
    return exception;
  }

  private static String location(String sheetName, String address) {
    return sheetName + "!" + address;
  }

  private static String functionLabel(String formula) {
    String functionName = leadingFunctionName(formula);
    return functionName == null ? "" : " function " + functionName;
  }

  private static String leadingFunctionName(String formula) {
    if (formula == null) {
      return null;
    }
    int openParenthesis = formula.indexOf('(');
    if (openParenthesis <= 0) {
      return null;
    }
    String candidate = formula.substring(0, openParenthesis).trim();
    if (candidate.isEmpty()) {
      return null;
    }
    for (int index = 0; index < candidate.length(); index++) {
      char current = candidate.charAt(index);
      if (!Character.isLetter(current) && current != '_' && current != '.') {
        return null;
      }
    }
    return candidate;
  }
}
