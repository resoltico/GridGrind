package dev.erst.gridgrind.engine.formulas;

import dev.erst.gridgrind.core.model.formula.ResolvedFormulaEnvironment;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Optional;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.ss.formula.atp.AnalysisToolPak;
import org.apache.poi.ss.formula.function.FunctionMetadataRegistry;
import org.jspecify.annotations.Nullable;

/**
 * Translates raw Apache POI formula evaluation exceptions into typed GridGrind exception hierarchy
 * entries.
 */
public final class FormulaExceptions {
  private static final Set<String> AUTHORED_UNSUPPORTED_CONSTRUCTS = Set.of("LAMBDA", "LET");
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
  public static RuntimeException wrap(
      String sheetName, String address, @Nullable String formula, RuntimeException exception) {
    return wrap(
        FormulaContext.from(ResolvedFormulaEnvironment.defaults()),
        sheetName,
        address,
        formula,
        exception);
  }

  /** Translates a native formula failure using the active evaluator context. */
  public static RuntimeException wrap(
      FormulaRuntime formulaRuntime,
      String sheetName,
      String address,
      @Nullable String formula,
      RuntimeException exception) {
    return wrap(formulaRuntime.context(), sheetName, address, formula, exception);
  }

  private static RuntimeException wrap(
      FormulaContext context,
      String sheetName,
      String address,
      @Nullable String formula,
      RuntimeException exception) {
    String normalizedFormula = formula == null ? "" : formula;
    if (isMissingExternalWorkbookFailure(exception)) {
      String workbookName = missingExternalWorkbookName(exception, normalizedFormula);
      return new MissingExternalWorkbookException(
          sheetName,
          address,
          normalizedFormula,
          workbookName,
          "Missing external workbook"
              + workbookLabel(workbookName)
              + " at "
              + location(sheetName, address)
              + ": "
              + normalizedFormula,
          exception);
    }
    if (isUnregisteredUserDefinedFunctionFailure(context, exception, normalizedFormula)) {
      String functionName = leadingFunctionName(normalizedFormula).orElse(null);
      return new UnregisteredUserDefinedFunctionException(
          sheetName,
          address,
          normalizedFormula,
          functionName,
          "User-defined function "
              + functionName
              + " is not registered at "
              + location(sheetName, address)
              + ": "
              + normalizedFormula,
          exception);
    }
    if (isUnsupportedFormulaFailure(exception)) {
      return new UnsupportedFormulaException(
          sheetName,
          address,
          normalizedFormula,
          "Unsupported formula"
              + functionLabel(normalizedFormula)
              + " at "
              + location(sheetName, address)
              + ": "
              + normalizedFormula,
          exception);
    }
    if (isUnhandledFormulaEvaluationFailure(exception)) {
      return exception;
    }
    if (isUnsupportedAuthoredConstructFailure(normalizedFormula, exception)) {
      return new UnsupportedFormulaConstructException(
          sheetName,
          address,
          normalizedFormula,
          "Unsupported formula construct"
              + functionLabel(normalizedFormula)
              + " at "
              + location(sheetName, address)
              + ": "
              + normalizedFormula,
          exception);
    }
    if (isInvalidFormulaFailure(exception)) {
      return new InvalidFormulaException(
          sheetName,
          address,
          normalizedFormula,
          "Invalid formula at " + location(sheetName, address) + ": " + normalizedFormula,
          exception);
    }
    return exception;
  }

  private static boolean isUnsupportedAuthoredConstructFailure(
      String formula, Throwable exception) {
    return isInvalidFormulaFailure(exception)
        && unsupportedAuthoredConstructName(formula).isPresent();
  }

  private static boolean isUnsupportedFormulaFailure(Throwable exception) {
    return hasTypeNameContaining(exception, ".ss.formula.eval.")
        && hasTypeNameContaining(exception, "NotImplemented");
  }

  /** Recognizes a missing external workbook in the native cause chain. */
  public static boolean isMissingExternalWorkbookFailure(Throwable exception) {
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
    String functionName = leadingFunctionName(formula).orElse(null);
    return functionName == null ? "" : " function " + functionName;
  }

  private static String workbookLabel(@Nullable String workbookName) {
    return workbookName == null ? "" : " " + workbookName;
  }

  /** Extracts the missing workbook name or a formula reference fallback. */
  public static @Nullable String missingExternalWorkbookName(Throwable exception, String formula) {
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

  /** Recognizes an unregistered authored function under the active evaluator context. */
  public static boolean isUnregisteredUserDefinedFunctionFailure(
      FormulaContext context, Throwable exception, String formula) {
    if (!isUnsupportedFormulaFailure(exception)) {
      return false;
    }
    String functionName = leadingFunctionName(formula).orElse(null);
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

  /** Checks native built-in functions and explicitly recognized authored constructs. */
  public static boolean isKnownBuiltinFunction(String normalizedFunctionName) {
    String canonicalFunctionName = canonicalFunctionName(normalizedFunctionName);
    return AUTHORED_UNSUPPORTED_CONSTRUCTS.contains(canonicalFunctionName)
        || FunctionMetadataRegistry.getFunctionByName(canonicalFunctionName) != null
        || AnalysisToolPak.getSupportedFunctionNames().contains(canonicalFunctionName)
        || AnalysisToolPak.getNotSupportedFunctionNames().contains(canonicalFunctionName);
  }

  private static Optional<String> unsupportedAuthoredConstructName(String formula) {
    return leadingFunctionName(formula)
        .map(name -> name.toUpperCase(Locale.ROOT))
        .map(FormulaExceptions::canonicalFunctionName)
        .filter(AUTHORED_UNSUPPORTED_CONSTRUCTS::contains);
  }

  private static String canonicalFunctionName(String functionName) {
    return functionName.startsWith("_XLFN.")
        ? functionName.substring("_XLFN.".length())
        : functionName;
  }

  /** Checks whether the cause chain names the normalized function. */
  public static boolean messageMentionsFunctionName(
      Throwable exception, String normalizedFunctionName) {
    for (Throwable current = exception; current != null; current = current.getCause()) {
      String message = current.getMessage();
      if (message != null && message.toUpperCase(Locale.ROOT).contains(normalizedFunctionName)) {
        return true;
      }
    }
    return false;
  }

  /** Reads a leading function identifier from the formula expression. */
  public static Optional<String> leadingFunctionName(String formula) {
    if (formula == null) {
      return Optional.empty();
    }
    int openParenthesis = formula.indexOf('(');
    if (openParenthesis <= 0) {
      return Optional.empty();
    }
    String candidate = formula.substring(0, openParenthesis).trim();
    if (candidate.isEmpty()) {
      return Optional.empty();
    }
    for (int index = 0; index < candidate.length(); index++) {
      char current = candidate.charAt(index);
      if (!Character.isLetter(current) && current != '_' && current != '.') {
        return Optional.empty();
      }
    }
    return Optional.of(candidate);
  }

  /** Reads external workbook names in first-reference order. */
  public static List<String> externalWorkbookNames(String formula) {
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
