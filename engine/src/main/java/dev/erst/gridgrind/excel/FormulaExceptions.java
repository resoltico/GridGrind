package dev.erst.gridgrind.excel;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.formula.atp.AnalysisToolPak;
import org.apache.poi.ss.formula.eval.FunctionEval;
import org.apache.poi.ss.formula.function.FunctionMetadataRegistry;

/**
 * Translates raw Apache POI formula evaluation exceptions into typed GridGrind exception hierarchy
 * entries.
 */
final class FormulaExceptions {
  private static final Pattern EXTERNAL_WORKBOOK_NAME_PATTERN =
      Pattern.compile("Could not resolve external workbook name '([^']+)'");
  private static final Pattern EXTERNAL_WORKBOOK_REFERENCE_PATTERN =
      Pattern.compile("\\[([^\\[\\]]+)]");

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
    return wrap(
        ExcelFormulaEnvironment.defaults().runtimeContext(),
        sheetName,
        address,
        formula,
        exception);
  }

  static RuntimeException wrap(
      ExcelFormulaRuntime formulaRuntime,
      String sheetName,
      String address,
      String formula,
      RuntimeException exception) {
    return wrap(formulaRuntime.context(), sheetName, address, formula, exception);
  }

  private static RuntimeException wrap(
      ExcelFormulaRuntimeContext context,
      String sheetName,
      String address,
      String formula,
      RuntimeException exception) {
    if (isMissingExternalWorkbookFailure(exception)) {
      String workbookName = missingExternalWorkbookName(exception, formula);
      return new MissingExternalWorkbookException(
          sheetName,
          address,
          formula,
          workbookName,
          "Missing external workbook"
              + workbookLabel(workbookName)
              + " at "
              + location(sheetName, address)
              + ": "
              + formula,
          exception);
    }
    if (isUnregisteredUserDefinedFunctionFailure(context, exception, formula)) {
      String functionName = leadingFunctionName(formula);
      return new UnregisteredUserDefinedFunctionException(
          sheetName,
          address,
          formula,
          functionName,
          "User-defined function "
              + functionName
              + " is not registered at "
              + location(sheetName, address)
              + ": "
              + formula,
          exception);
    }
    if (isUnsupportedFormulaFailure(exception)) {
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
    if (isUnhandledFormulaEvaluationFailure(exception)) {
      return exception;
    }
    if (isInvalidFormulaFailure(exception)) {
      return new InvalidFormulaException(
          sheetName,
          address,
          formula,
          "Invalid formula at " + location(sheetName, address) + ": " + formula,
          exception);
    }
    return exception;
  }

  private static boolean isUnsupportedFormulaFailure(Throwable exception) {
    return hasTypeNameContaining(exception, ".ss.formula.eval.")
        && hasTypeNameContaining(exception, "NotImplemented");
  }

  static boolean isMissingExternalWorkbookFailure(Throwable exception) {
    return hasTypeNameContaining(exception, "WorkbookNotFoundException");
  }

  private static boolean isUnhandledFormulaEvaluationFailure(Throwable exception) {
    return hasTypeNameContaining(exception, ".ss.formula.eval.");
  }

  private static boolean isInvalidFormulaFailure(Throwable exception) {
    return hasTypeNameStartingWith(exception, "org.apache.poi.ss.formula.")
        || hasStackFrameStartingWith(exception, "org.apache.poi.ss.formula.");
  }

  private static boolean hasTypeNameContaining(Throwable exception, String fragment) {
    for (Throwable current = exception; current != null; current = current.getCause()) {
      if (current.getClass().getName().contains(fragment)) {
        return true;
      }
    }
    return false;
  }

  private static boolean hasTypeNameStartingWith(Throwable exception, String prefix) {
    for (Throwable current = exception; current != null; current = current.getCause()) {
      if (current.getClass().getName().startsWith(prefix)) {
        return true;
      }
    }
    return false;
  }

  private static boolean hasStackFrameStartingWith(Throwable exception, String prefix) {
    for (Throwable current = exception; current != null; current = current.getCause()) {
      for (StackTraceElement frame : current.getStackTrace()) {
        if (frame.getClassName().startsWith(prefix)) {
          return true;
        }
      }
    }
    return false;
  }

  private static String location(String sheetName, String address) {
    return sheetName + "!" + address;
  }

  private static String functionLabel(String formula) {
    String functionName = leadingFunctionName(formula);
    return functionName == null ? "" : " function " + functionName;
  }

  private static String workbookLabel(String workbookName) {
    return workbookName == null ? "" : " " + workbookName;
  }

  static String missingExternalWorkbookName(Throwable exception, String formula) {
    for (Throwable current = exception; current != null; current = current.getCause()) {
      String message = current.getMessage();
      if (message == null) {
        continue;
      }
      Matcher matcher = EXTERNAL_WORKBOOK_NAME_PATTERN.matcher(message);
      if (matcher.find()) {
        return matcher.group(1);
      }
    }
    List<String> referencedNames = externalWorkbookNames(formula);
    return referencedNames.isEmpty() ? null : referencedNames.getFirst();
  }

  static boolean isUnregisteredUserDefinedFunctionFailure(
      ExcelFormulaRuntimeContext context, Throwable exception, String formula) {
    if (!isUnsupportedFormulaFailure(exception)) {
      return false;
    }
    String functionName = leadingFunctionName(formula);
    if (functionName == null) {
      return false;
    }
    String normalized = functionName.toUpperCase(Locale.ROOT);
    if (isKnownBuiltinFunction(normalized)) {
      return false;
    }
    return !context.hasUserDefinedFunction(normalized)
        && messageMentionsFunctionName(exception, normalized);
  }

  private static boolean isKnownBuiltinFunction(String normalizedFunctionName) {
    return FunctionMetadataRegistry.getFunctionByName(normalizedFunctionName) != null
        || FunctionEval.getSupportedFunctionNames().contains(normalizedFunctionName)
        || FunctionEval.getNotSupportedFunctionNames().contains(normalizedFunctionName)
        || AnalysisToolPak.getSupportedFunctionNames().contains(normalizedFunctionName)
        || AnalysisToolPak.getNotSupportedFunctionNames().contains(normalizedFunctionName);
  }

  private static boolean messageMentionsFunctionName(
      Throwable exception, String normalizedFunctionName) {
    for (Throwable current = exception; current != null; current = current.getCause()) {
      String message = current.getMessage();
      if (message != null && message.toUpperCase(Locale.ROOT).contains(normalizedFunctionName)) {
        return true;
      }
    }
    return false;
  }

  static String leadingFunctionName(String formula) {
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

  static List<String> externalWorkbookNames(String formula) {
    if (formula == null || formula.isBlank()) {
      return List.of();
    }
    Matcher matcher = EXTERNAL_WORKBOOK_REFERENCE_PATTERN.matcher(formula);
    Set<String> names = new LinkedHashSet<>();
    while (matcher.find()) {
      names.add(matcher.group(1));
    }
    return List.copyOf(names);
  }
}
