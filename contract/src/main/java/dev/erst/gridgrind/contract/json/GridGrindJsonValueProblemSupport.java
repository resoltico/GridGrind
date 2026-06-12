package dev.erst.gridgrind.contract.json;

import java.util.List;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.jspecify.annotations.Nullable;
import tools.jackson.core.JacksonException;
import tools.jackson.databind.exc.InvalidFormatException;

/** Owns non-subtype public wording for JSON value-shape failures. */
final class GridGrindJsonValueProblemSupport {
  private static final Pattern MISSING_REQUIRED_FIELD_PATTERN =
      Pattern.compile("missing required creator property '([^']+)'", Pattern.CASE_INSENSITIVE);
  private static final Pattern MISSING_TYPE_ID_FIELD_PATTERN =
      Pattern.compile("missing type id property '([^']+)'", Pattern.CASE_INSENSITIVE);
  private static final Pattern NULL_FIELD_PROBLEM_PATTERN =
      Pattern.compile("([A-Za-z0-9.\\[\\]_]+) must not be null", Pattern.CASE_INSENSITIVE);

  private GridGrindJsonValueProblemSupport() {}

  static String mismatchedInputMessage(
      tools.jackson.databind.exc.MismatchedInputException exception) {
    String original = exception.getOriginalMessage();
    if (original != null && original.contains("Floating-point value")) {
      return floatingPointIntoIntegerMessage(exception);
    }
    return productOwnedJacksonMessage(
        GridGrindJsonProblemMessageSupport.cleanJacksonMessage(original));
  }

  static String enumValueMessage(InvalidFormatException exception) {
    Object value = exception.getValue();
    String renderedValue = value == null ? "null" : value.toString();
    String fieldName =
        Optional.ofNullable(exception.getPath().isEmpty() ? null : exception.getPath().getLast())
            .map(JacksonException.Reference::getPropertyName)
            .filter(GridGrindJsonValueProblemSupport::hasNonBlankFieldName)
            .orElse(null);
    String allowedValues =
        java.util.Arrays.stream(exception.getTargetType().getEnumConstants())
            .map(constant -> ((Enum<?>) constant).name())
            .collect(java.util.stream.Collectors.joining(", "));
    if (fieldName != null) {
      return "Unsupported value '"
          + renderedValue
          + "' for field '"
          + fieldName
          + "'; expected one of: "
          + allowedValues;
    }
    return "Unsupported value '" + renderedValue + "'; expected one of: " + allowedValues;
  }

  static boolean hasNonBlankFieldName(@Nullable String fieldName) {
    return fieldName != null && !fieldName.isBlank();
  }

  static String productOwnedJacksonMessage(@Nullable String cleaned) {
    String normalized = GridGrindJsonProblemMessageSupport.cleanJacksonMessage(cleaned);
    Matcher missingRequiredField = MISSING_REQUIRED_FIELD_PATTERN.matcher(normalized);
    if (missingRequiredField.find()) {
      return "Missing required field '" + missingRequiredField.group(1) + "'";
    }
    Matcher missingTypeIdField = MISSING_TYPE_ID_FIELD_PATTERN.matcher(normalized);
    if (missingTypeIdField.find()) {
      return "Missing required field '" + missingTypeIdField.group(1) + "'";
    }
    Matcher nullFieldProblem = NULL_FIELD_PROBLEM_PATTERN.matcher(normalized);
    if (nullFieldProblem.find()) {
      return "Missing required field '" + nullFieldProblem.group(1) + "'";
    }
    if (normalized.startsWith("Cannot deserialize value")) {
      return "JSON value has the wrong shape for this field";
    }
    if (normalized.startsWith("Cannot construct instance of")) {
      return "JSON object is missing required fields or has the wrong shape";
    }
    return normalized;
  }

  private static String floatingPointIntoIntegerMessage(
      tools.jackson.databind.exc.MismatchedInputException exception) {
    List<JacksonException.Reference> path = exception.getPath();
    if (path.isEmpty()) {
      return "JSON value must be an integer value";
    }
    String propertyName = path.getLast().getPropertyName();
    if (propertyName != null) {
      return "Field '" + propertyName + "' must be an integer value";
    }
    return "JSON value at '"
        + GridGrindJsonPayloadMetadataSupport.renderPath(path)
        + "' must be an integer value";
  }
}
