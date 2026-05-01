package dev.erst.gridgrind.contract.json;

import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import tools.jackson.core.JacksonException;
import tools.jackson.core.TokenStreamLocation;
import tools.jackson.core.exc.StreamConstraintsException;
import tools.jackson.core.exc.StreamReadException;
import tools.jackson.databind.DatabindException;

/** Owns public error wording and payload-location extraction for JSON failures. */
final class GridGrindJsonMessageSupport {
  private static final Pattern MISSING_REQUIRED_FIELD_PATTERN =
      Pattern.compile("missing required creator property '([^']+)'", Pattern.CASE_INSENSITIVE);
  private static final Pattern MISSING_TYPE_ID_FIELD_PATTERN =
      Pattern.compile("missing type id property '([^']+)'", Pattern.CASE_INSENSITIVE);
  private static final Pattern NULL_FIELD_PROBLEM_PATTERN =
      Pattern.compile(
          "(?:problem: )?([A-Za-z0-9.\\[\\]_]+) must not be null", Pattern.CASE_INSENSITIVE);

  private GridGrindJsonMessageSupport() {}

  static IllegalArgumentException invalidPayload(JacksonException exception) {
    PayloadMetadata metadata = payloadMetadata(exception);
    Optional<Throwable> validationCause = validationCause(exception);
    if (exception instanceof StreamReadException) {
      return new InvalidJsonException(
          message(exception),
          metadata.jsonPath().orElse(null),
          metadata.jsonLine().orElse(null),
          metadata.jsonColumn().orElse(null),
          exception);
    }
    if (validationCause.isPresent()) {
      return new InvalidRequestException(
          message(validationCause.orElseThrow()),
          metadata.jsonPath().orElse(null),
          metadata.jsonLine().orElse(null),
          metadata.jsonColumn().orElse(null),
          exception);
    }
    if (exception instanceof DatabindException) {
      return new InvalidRequestShapeException(
          message(exception),
          metadata.jsonPath().orElse(null),
          metadata.jsonLine().orElse(null),
          metadata.jsonColumn().orElse(null),
          exception);
    }
    return new InvalidJsonException(
        message(exception),
        metadata.jsonPath().orElse(null),
        metadata.jsonLine().orElse(null),
        metadata.jsonColumn().orElse(null),
        exception);
  }

  static IllegalArgumentException invalidRequestPayload(JacksonException exception) {
    if (exception instanceof StreamConstraintsException) {
      return GridGrindJsonMapperSupport.requestTooLarge(exception);
    }
    return invalidPayload(exception);
  }

  static String message(Throwable throwable) {
    if (throwable
        instanceof tools.jackson.databind.exc.InvalidTypeIdException invalidTypeIdException) {
      return unknownTypeValueMessage(invalidTypeIdException);
    }
    if (throwable
        instanceof
        tools.jackson.databind.exc.UnrecognizedPropertyException unrecognizedPropertyException) {
      return "Unknown field '" + unrecognizedPropertyException.getPropertyName() + "'";
    }
    if (throwable
        instanceof tools.jackson.databind.exc.MismatchedInputException mismatchedInputException) {
      return mismatchedInputMessage(mismatchedInputException);
    }
    String message =
        throwable instanceof JacksonException jacksonException
            ? jacksonException.getOriginalMessage()
            : throwable.getMessage();
    return productOwnedJacksonMessage(cleanJacksonMessage(message));
  }

  static String mismatchedInputMessage(
      tools.jackson.databind.exc.MismatchedInputException exception) {
    String original = exception.getOriginalMessage();
    if (original != null && original.startsWith("Cannot map `null` into type")) {
      return nullIntoPrimitiveMessage(exception);
    }
    if (original != null && original.contains("Floating-point value")) {
      return floatingPointIntoIntegerMessage(exception);
    }
    return productOwnedJacksonMessage(cleanJacksonMessage(original));
  }

  static String cleanJacksonMessage(String message) {
    if (message == null || message.isBlank()) {
      return "Invalid JSON payload";
    }
    int startMarkerIndex = message.indexOf(" (start marker at [Source:");
    String trimmed = startMarkerIndex >= 0 ? message.substring(0, startMarkerIndex) : message;
    String cleaned =
        trimmed
            .replaceAll(" as a subtype of `[^`]+`", "")
            .replaceAll(" \\(for POJO property '[^']+'\\)", "")
            .replaceAll(" \\(set [^)]*to allow\\)", "")
            .strip();
    return cleaned.isBlank() ? "Invalid JSON payload" : cleaned;
  }

  static Optional<Integer> jsonLine(TokenStreamLocation location) {
    if (location == null) {
      return Optional.empty();
    }
    int line = location.getLineNr();
    return line > 0 ? Optional.of(line) : Optional.empty();
  }

  static Optional<Integer> jsonColumn(TokenStreamLocation location) {
    if (location == null) {
      return Optional.empty();
    }
    int column = location.getColumnNr();
    return column > 0 ? Optional.of(column) : Optional.empty();
  }

  private static Optional<Throwable> validationCause(Throwable throwable) {
    Throwable current = throwable;
    while (current != null) {
      if (current instanceof IllegalArgumentException
          || current instanceof NullPointerException
          || current instanceof java.time.DateTimeException) {
        return Optional.of(current);
      }
      current = current.getCause();
    }
    return Optional.empty();
  }

  private static String unknownTypeValueMessage(
      tools.jackson.databind.exc.InvalidTypeIdException exception) {
    String typeId = exception.getTypeId();
    if (typeId == null) {
      return productOwnedJacksonMessage(cleanJacksonMessage(exception.getOriginalMessage()));
    }
    String defaultMessage = "Unknown type value '" + typeId + "'";
    Optional<String> containerName = terminalContainerName(exception.getPath());
    if (containerName.isPresent() && "source".equals(containerName.orElseThrow())) {
      if ("FILE".equals(typeId)) {
        return defaultMessage
            + "; use source.type='EXISTING' to open a workbook from disk"
            + " (FILE is only valid for source-backed authored payload inputs)";
      }
      return defaultMessage;
    }
    if (containerName.isPresent() && "assertion".equals(containerName.orElseThrow())) {
      return switch (typeId) {
        case "EXPECT_PRESENT" ->
            defaultMessage
                + "; use one explicit family assertion such as EXPECT_NAMED_RANGE_PRESENT,"
                + " EXPECT_TABLE_PRESENT, EXPECT_PIVOT_TABLE_PRESENT, or EXPECT_CHART_PRESENT";
        case "EXPECT_ABSENT" ->
            defaultMessage
                + "; use one explicit family assertion such as EXPECT_NAMED_RANGE_ABSENT,"
                + " EXPECT_TABLE_ABSENT, EXPECT_PIVOT_TABLE_ABSENT, or EXPECT_CHART_ABSENT";
        default -> defaultMessage;
      };
    }
    return defaultMessage;
  }

  private static Optional<String> terminalContainerName(List<JacksonException.Reference> path) {
    if (path.isEmpty()) {
      return Optional.empty();
    }
    return Optional.ofNullable(path.getLast().getPropertyName());
  }

  private static String nullIntoPrimitiveMessage(
      tools.jackson.databind.exc.MismatchedInputException exception) {
    List<JacksonException.Reference> path = exception.getPath();
    if (path.isEmpty()) {
      return "Missing required field";
    }
    String propertyName = path.getLast().getPropertyName();
    if (propertyName == null) {
      return "Missing required field";
    }
    return "Missing required field '" + propertyName + "'";
  }

  private static String floatingPointIntoIntegerMessage(
      tools.jackson.databind.exc.MismatchedInputException exception) {
    List<JacksonException.Reference> path = exception.getPath();
    if (path.isEmpty()) {
      return "JSON value must be an integer value";
    }
    String propertyName = path.getLast().getPropertyName();
    if (propertyName == null) {
      return "JSON value must be an integer value";
    }
    return "Field '" + propertyName + "' must be an integer value";
  }

  private static String productOwnedJacksonMessage(String cleaned) {
    Matcher missingRequiredField = MISSING_REQUIRED_FIELD_PATTERN.matcher(cleaned);
    if (missingRequiredField.find()) {
      return "Missing required field '" + missingRequiredField.group(1) + "'";
    }
    Matcher missingTypeIdField = MISSING_TYPE_ID_FIELD_PATTERN.matcher(cleaned);
    if (missingTypeIdField.find()) {
      return "Missing required field '" + missingTypeIdField.group(1) + "'";
    }
    Matcher nullFieldProblem = NULL_FIELD_PROBLEM_PATTERN.matcher(cleaned);
    if (nullFieldProblem.find()) {
      return "Missing required field '" + nullFieldProblem.group(1) + "'";
    }
    if (cleaned.startsWith("Cannot deserialize value")) {
      return "JSON value has the wrong shape for this field";
    }
    if (cleaned.startsWith("Cannot construct instance of")) {
      return "JSON object is missing required fields or has the wrong shape";
    }
    return cleaned;
  }

  private static PayloadMetadata payloadMetadata(JacksonException exception) {
    return new PayloadMetadata(
        Optional.ofNullable(jsonPath(exception)),
        jsonLine(exception.getLocation()),
        jsonColumn(exception.getLocation()));
  }

  private static String jsonPath(JacksonException exception) {
    StringBuilder path = new StringBuilder();
    for (JacksonException.Reference reference : exception.getPath()) {
      String propertyName = reference.getPropertyName();
      if (propertyName != null) {
        if (!path.isEmpty()) {
          path.append('.');
        }
        path.append(propertyName);
        continue;
      }
      path.append('[').append(reference.getIndex()).append(']');
    }
    return path.isEmpty() ? null : path.toString();
  }

  private record PayloadMetadata(
      Optional<String> jsonPath, Optional<Integer> jsonLine, Optional<Integer> jsonColumn) {
    private PayloadMetadata {
      jsonPath = Objects.requireNonNullElseGet(jsonPath, Optional::empty);
      jsonLine = Objects.requireNonNullElseGet(jsonLine, Optional::empty);
      jsonColumn = Objects.requireNonNullElseGet(jsonColumn, Optional::empty);
    }
  }
}
