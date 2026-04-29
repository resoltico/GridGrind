package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/** One template-backed user-defined function registered for formula evaluation. */
public record FormulaUdfFunctionInput(
    String name,
    Integer minimumArgumentCount,
    Integer maximumArgumentCount,
    String formulaTemplate) {
  private static final Pattern FUNCTION_NAME_PATTERN = Pattern.compile("[A-Za-z_][A-Za-z0-9_.]*");
  private static final Pattern ARGUMENT_PLACEHOLDER_PATTERN =
      Pattern.compile("\\bARG([1-9][0-9]*)\\b", Pattern.CASE_INSENSITIVE);

  public FormulaUdfFunctionInput {
    name = requireFunctionName(name);
    minimumArgumentCount = requireMinimumArgumentCount(minimumArgumentCount);
    Objects.requireNonNull(maximumArgumentCount, "maximumArgumentCount must not be null");
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

  private static Integer requireMinimumArgumentCount(Integer minimumArgumentCount) {
    Objects.requireNonNull(minimumArgumentCount, "minimumArgumentCount must not be null");
    if (minimumArgumentCount < 0) {
      throw new IllegalArgumentException("minimumArgumentCount must not be negative");
    }
    return minimumArgumentCount;
  }

  private static void requireMaximumArgumentCount(
      Integer minimumArgumentCount, Integer maximumArgumentCount) {
    Objects.requireNonNull(maximumArgumentCount, "maximumArgumentCount must not be null");
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

  @JsonCreator
  static FormulaUdfFunctionInput create(
      @JsonProperty("name") String name,
      @JsonProperty("minimumArgumentCount") Integer minimumArgumentCount,
      @JsonProperty("maximumArgumentCount") Integer maximumArgumentCount,
      @JsonProperty("formulaTemplate") String formulaTemplate) {
    Integer normalizedMinimum = requireMinimumArgumentCount(minimumArgumentCount);
    return new FormulaUdfFunctionInput(
        name,
        normalizedMinimum,
        maximumArgumentCount == null ? normalizedMinimum : maximumArgumentCount,
        formulaTemplate);
  }
}
