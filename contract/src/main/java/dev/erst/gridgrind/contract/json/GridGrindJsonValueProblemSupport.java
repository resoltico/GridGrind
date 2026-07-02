package dev.erst.gridgrind.contract.json;

import java.util.List;
import java.util.Optional;
import org.jspecify.annotations.Nullable;
import tools.jackson.core.JacksonException;
import tools.jackson.core.JsonToken;
import tools.jackson.databind.exc.InvalidFormatException;
import tools.jackson.databind.exc.MismatchedInputException;

/** Owns non-subtype public wording for JSON value-shape failures. */
final class GridGrindJsonValueProblemSupport {
  private GridGrindJsonValueProblemSupport() {}

  static String mismatchedInputMessage(MismatchedInputException exception) {
    if (isFloatingPointIntoInteger(exception)) {
      return floatingPointIntoIntegerMessage(exception);
    }
    String original = exception.getOriginalMessage();
    if (original == null || original.isBlank()) {
      return GridGrindJsonProblemMessageSupport.cleanJacksonMessage(original);
    }
    return exception.getCurrentToken() == JsonToken.START_OBJECT
        ? genericObjectShapeMessage()
        : genericValueShapeMessage();
  }

  static boolean isFloatingPointIntoInteger(MismatchedInputException exception) {
    if (!integralTargetType(exception.getTargetType())) {
      return false;
    }
    if (exception.getCurrentToken() == JsonToken.VALUE_NUMBER_FLOAT) {
      return true;
    }
    String originalMessage = exception.getOriginalMessage();
    return originalMessage != null && originalMessage.contains("Floating-point value");
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

  static String genericObjectShapeMessage() {
    return "JSON object is missing required fields or has the wrong shape";
  }

  static String genericValueShapeMessage() {
    return "JSON value has the wrong shape for this field";
  }

  private static String floatingPointIntoIntegerMessage(MismatchedInputException exception) {
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

  private static boolean integralTargetType(@Nullable Class<?> targetType) {
    return targetType == byte.class
        || targetType == short.class
        || targetType == int.class
        || targetType == long.class
        || targetType == Byte.class
        || targetType == Short.class
        || targetType == Integer.class
        || targetType == Long.class
        || targetType == java.math.BigInteger.class;
  }
}
