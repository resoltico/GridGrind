package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** One template-backed user-defined function registered for formula evaluation. */
@CatalogType(
    id = "FormulaUdfFunctionSpec",
    summary =
        "One template-backed user-defined function. formulaTemplate may reference ARG1, ARG2, and higher placeholders. maximumArgumentCount defaults to minimumArgumentCount when omitted.",
    displayOrder = 28)
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_UDF_FORMULA_TEMPLATE_INJECTION_DEFENSE_IN_DEPTH
})
public record FormulaUdfFunctionSpec(
    String name, int minimumArgumentCount, int maximumArgumentCount, String formulaTemplate) {
  private static final Pattern FUNCTION_NAME_PATTERN = Pattern.compile("[A-Za-z_][A-Za-z0-9_.]*");
  private static final Pattern ARGUMENT_PLACEHOLDER_PATTERN =
      Pattern.compile("\\bARG([1-9][0-9]*)\\b", Pattern.CASE_INSENSITIVE);

  public FormulaUdfFunctionSpec {
    name = requireFunctionName(name);
    requireMinimumArgumentCount(minimumArgumentCount);
    requireMaximumArgumentCount(minimumArgumentCount, maximumArgumentCount);
    formulaTemplate = normalizeFormulaTemplate(formulaTemplate);
    FormulaInputSecurity.rejectDde(formulaTemplate);
    int highestPlaceholderIndex = highestPlaceholderIndex(formulaTemplate);
    if (highestPlaceholderIndex > maximumArgumentCount) {
      throw new IllegalArgumentException(
          "formulaTemplate references ARG"
              + highestPlaceholderIndex
              + " but maximumArgumentCount is "
              + maximumArgumentCount);
    }
  }

  /** Creates a function whose minimum and maximum arity are the same exact count. */
  public FormulaUdfFunctionSpec(String name, int exactArgumentCount, String formulaTemplate) {
    this(name, exactArgumentCount, exactArgumentCount, formulaTemplate);
  }

  private static String requireFunctionName(String name) {
    TextValidation.requireNonBlank(name, "name");
    if (!FUNCTION_NAME_PATTERN.matcher(name).matches()) {
      throw new IllegalArgumentException(
          "name must start with a letter or underscore and contain only letters, digits, underscores, or dots");
    }
    return name;
  }

  private static void requireMinimumArgumentCount(int minimumArgumentCount) {
    if (minimumArgumentCount < 0) {
      throw new IllegalArgumentException("minimumArgumentCount must not be negative");
    }
  }

  private static void requireMaximumArgumentCount(
      int minimumArgumentCount, int maximumArgumentCount) {
    if (maximumArgumentCount < minimumArgumentCount) {
      throw new IllegalArgumentException(
          "maximumArgumentCount must not be less than minimumArgumentCount");
    }
  }

  private static String normalizeFormulaTemplate(String formulaTemplate) {
    Objects.requireNonNull(formulaTemplate, "formulaTemplate must not be null");
    String trimmed = formulaTemplate.trim();
    TextValidation.requireNonBlank(trimmed, "formulaTemplate");
    return trimmed.startsWith("=") ? trimmed.substring(1).trim() : trimmed;
  }

  private static int highestPlaceholderIndex(String formulaTemplate) {
    Matcher matcher = ARGUMENT_PLACEHOLDER_PATTERN.matcher(formulaTemplate);
    int highest = 0;
    while (matcher.find()) {
      highest = Math.max(highest, Integer.parseInt(matcher.group(1)));
    }
    return highest;
  }
}
