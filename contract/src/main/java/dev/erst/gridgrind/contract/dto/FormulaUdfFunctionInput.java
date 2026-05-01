package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** One template-backed user-defined function registered for formula evaluation. */
public record FormulaUdfFunctionInput(
    String name, int minimumArgumentCount, int maximumArgumentCount, String formulaTemplate) {
  private static final Pattern FUNCTION_NAME_PATTERN = Pattern.compile("[A-Za-z_][A-Za-z0-9_.]*");
  private static final Pattern ARGUMENT_PLACEHOLDER_PATTERN =
      Pattern.compile("\\bARG([1-9][0-9]*)\\b", Pattern.CASE_INSENSITIVE);

  public FormulaUdfFunctionInput {
    name = requireFunctionName(name);
    requireMinimumArgumentCount(minimumArgumentCount);
    requireMaximumArgumentCount(minimumArgumentCount, maximumArgumentCount);
    formulaTemplate = normalizeFormulaTemplate(formulaTemplate);
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
  public FormulaUdfFunctionInput(String name, int exactArgumentCount, String formulaTemplate) {
    this(name, exactArgumentCount, exactArgumentCount, formulaTemplate);
  }

  /** Creates a UDF definition with explicit minimum and maximum arity. */
  @JsonCreator
  public FormulaUdfFunctionInput(
      @JsonProperty("name") String name,
      @JsonProperty("minimumArgumentCount") int minimumArgumentCount,
      @JsonProperty("maximumArgumentCount") Integer maximumArgumentCount,
      @JsonProperty("formulaTemplate") String formulaTemplate) {
    this(
        name,
        minimumArgumentCount,
        Objects.requireNonNull(maximumArgumentCount, "maximumArgumentCount must not be null")
            .intValue(),
        formulaTemplate);
  }

  private static String requireFunctionName(String name) {
    Objects.requireNonNull(name, "name must not be null");
    if (name.isBlank()) {
      throw new IllegalArgumentException("name must not be blank");
    }
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
    if (trimmed.isBlank()) {
      throw new IllegalArgumentException("formulaTemplate must not be blank");
    }
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
