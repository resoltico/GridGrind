package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.dto.GridGrindRequestProblemSupport;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import tools.jackson.core.JacksonException;
import tools.jackson.core.JacksonException.Reference;
import tools.jackson.core.JsonParser;
import tools.jackson.databind.exc.MismatchedInputException;

/** Derives precise nested request-field paths for step payload failures. */
final class WorkbookStepJsonFailurePathSupport {
  private static final Pattern FIELD_MESSAGE_PATTERN = Pattern.compile("^Field '([^']+)' must .*");
  private static final Pattern RAW_MISSING_REQUIRED_FIELD_PATTERN =
      Pattern.compile("missing required creator property '([^']+)'", Pattern.CASE_INSENSITIVE);
  private static final Pattern RAW_MISSING_TYPE_ID_FIELD_PATTERN =
      Pattern.compile("missing type id property '([^']+)'", Pattern.CASE_INSENSITIVE);

  private WorkbookStepJsonFailurePathSupport() {}

  static String qualifiedFieldName(String fieldName, Exception failure) {
    String nestedFieldPath = nestedFieldPath(failure);
    if (nestedFieldPath.isEmpty()) {
      return fieldName;
    }
    if (nestedFieldPath.equals(fieldName)
        || nestedFieldPath.startsWith(fieldName + ".")
        || nestedFieldPath.startsWith(fieldName + "[")) {
      return nestedFieldPath;
    }
    return nestedFieldPath.startsWith("[")
        ? fieldName + nestedFieldPath
        : fieldName + "." + nestedFieldPath;
  }

  static JacksonException wrapIllegalArgumentFailure(
      JsonParser parser, String fieldName, IllegalArgumentException exception) {
    MismatchedInputException failure =
        MismatchedInputException.from(
            parser,
            WorkbookStep.class,
            Objects.requireNonNullElse(exception.getMessage(), "Invalid request shape"));
    failure.initCause(exception);
    return failure.prependPath(WorkbookStep.class, qualifiedFieldName(fieldName, exception));
  }

  static JacksonException wrapJacksonFailure(String fieldName, JacksonException exception) {
    return exception.prependPath(
        WorkbookStep.class,
        exception.getPath().isEmpty() ? qualifiedFieldName(fieldName, exception) : fieldName);
  }

  private static String nestedFieldPath(Exception failure) {
    if (failure instanceof JacksonException jacksonException) {
      String renderedPath = renderPath(jacksonException.getPath());
      if (!renderedPath.isEmpty()) {
        return renderedPath;
      }
      String inferredFromOriginalMessage = inferFieldPath(jacksonException.getOriginalMessage());
      if (!inferredFromOriginalMessage.isEmpty()) {
        return inferredFromOriginalMessage;
      }
    }
    return inferFieldPath(Objects.requireNonNullElse(failure.getMessage(), ""));
  }

  private static String inferFieldPath(String message) {
    String normalized = Objects.requireNonNullElse(message, "").trim();
    if (normalized.isEmpty()) {
      return "";
    }
    var publicPath = GridGrindRequestProblemSupport.jsonPathFromMessage(normalized);
    if (publicPath.isPresent()) {
      return publicPath.orElseThrow();
    }
    Matcher fieldMessage = FIELD_MESSAGE_PATTERN.matcher(normalized);
    if (fieldMessage.matches()) {
      return fieldMessage.group(1);
    }
    Matcher missingRequiredField = RAW_MISSING_REQUIRED_FIELD_PATTERN.matcher(normalized);
    if (missingRequiredField.find()) {
      return missingRequiredField.group(1);
    }
    Matcher missingTypeIdField = RAW_MISSING_TYPE_ID_FIELD_PATTERN.matcher(normalized);
    if (missingTypeIdField.find()) {
      return missingTypeIdField.group(1);
    }
    return "";
  }

  private static String renderPath(java.util.List<Reference> path) {
    StringBuilder rendered = new StringBuilder();
    for (Reference reference : path) {
      String propertyName = reference.getPropertyName();
      if (propertyName != null) {
        if (!rendered.isEmpty()) {
          rendered.append('.');
        }
        rendered.append(propertyName);
      } else {
        rendered.append('[').append(reference.getIndex()).append(']');
      }
    }
    return rendered.toString();
  }
}
