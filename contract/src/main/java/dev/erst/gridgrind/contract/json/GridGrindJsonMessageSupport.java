package dev.erst.gridgrind.contract.json;

import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.jspecify.annotations.Nullable;
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
      Pattern.compile("([A-Za-z0-9.\\[\\]_]+) must not be null", Pattern.CASE_INSENSITIVE);

  private GridGrindJsonMessageSupport() {}

  static IllegalArgumentException invalidPayload(JacksonException exception) {
    PayloadMetadata metadata = payloadMetadata(exception);
    Optional<Throwable> validationCause = validationCause(exception);
    if (exception instanceof StreamReadException) {
      return new InvalidJsonException(
          message(exception),
          metadata.jsonPath(),
          metadata.jsonLine(),
          metadata.jsonColumn(),
          exception);
    }
    if (validationCause.isPresent()) {
      return new InvalidRequestException(
          message(validationCause.orElseThrow()),
          metadata.jsonPath(),
          metadata.jsonLine(),
          metadata.jsonColumn(),
          exception);
    }
    if (exception instanceof DatabindException) {
      return new InvalidRequestShapeException(
          message(exception),
          metadata.jsonPath(),
          metadata.jsonLine(),
          metadata.jsonColumn(),
          exception);
    }
    return new InvalidJsonException(
        message(exception),
        metadata.jsonPath(),
        metadata.jsonLine(),
        metadata.jsonColumn(),
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
    if (original != null && original.contains("Floating-point value")) {
      return floatingPointIntoIntegerMessage(exception);
    }
    return productOwnedJacksonMessage(cleanJacksonMessage(original));
  }

  static String cleanJacksonMessage(@Nullable String message) {
    if (message == null || message.isBlank()) {
      return "Invalid JSON payload";
    }
    int startMarkerIndex = message.indexOf(" (start marker at [Source:");
    String trimmed = startMarkerIndex >= 0 ? message.substring(0, startMarkerIndex) : message;
    String cleaned =
        trimmed
            .replaceAll(" as a subtype of `[^`]+`", "")
            .replaceAll(" \\(for POJO property '[^']+'\\)", "")
            .replaceAll(" \\(but could if coercion[^)]*\\)", "")
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
          || current instanceof java.time.DateTimeException) {
        return Optional.of(current);
      }
      if (current instanceof NullPointerException npe && isExplicitNullCheck(npe)) {
        return Optional.of(current);
      }
      current = current.getCause();
    }
    return Optional.empty();
  }

  private static boolean isExplicitNullCheck(NullPointerException npe) {
    String message = npe.getMessage();
    return message != null && message.endsWith("must not be null");
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
      return withCandidates(defaultMessage, exception, typeId);
    }
    if (containerName.isPresent() && "assertion".equals(containerName.orElseThrow())) {
      return switch (typeId) {
        case "EXPECT_PRESENT" ->
            legacyPresenceAssertionMessage(defaultMessage, exception, "_PRESENT");
        case "EXPECT_ABSENT" ->
            legacyPresenceAssertionMessage(defaultMessage, exception, "_ABSENT");
        default -> withCandidates(defaultMessage, exception, typeId);
      };
    }
    return withCandidates(defaultMessage, exception, typeId);
  }

  private static String legacyPresenceAssertionMessage(
      String base, tools.jackson.databind.exc.InvalidTypeIdException exception, String suffix) {
    List<String> candidates =
        GridGrindJsonSubtypeSupport.typeIds(exception.getBaseType().getRawClass()).stream()
            .filter(id -> id.endsWith(suffix))
            .toList();
    return base + "; use one explicit family assertion such as " + String.join(", ", candidates);
  }

  private static String withCandidates(
      String base, tools.jackson.databind.exc.InvalidTypeIdException exception, String typeId) {
    tools.jackson.databind.JavaType baseType = exception.getBaseType();
    if (baseType == null) {
      return base;
    }
    List<String> candidates = similarTypeIds(baseType.getRawClass(), typeId);
    if (candidates.isEmpty()) {
      return base;
    }
    return base + "; similar valid values: " + String.join(", ", candidates);
  }

  private static List<String> similarTypeIds(Class<?> baseClass, String typeId) {
    List<String> all = GridGrindJsonSubtypeSupport.typeIds(baseClass);
    String normalized = typeId.toUpperCase(java.util.Locale.ROOT);
    int threshold = Math.min(3, Math.max(1, normalized.length() / 5));
    return all.stream()
        .filter(id -> editDistance(normalized, id) <= threshold)
        .sorted(
            java.util.Comparator.<String>comparingInt(id -> editDistance(normalized, id))
                .thenComparing(java.util.Comparator.naturalOrder()))
        .limit(3)
        .toList();
  }

  @SuppressWarnings("PMD.AvoidArrayLoops")
  private static int editDistance(String a, String b) {
    int m = a.length();
    int n = b.length();
    int[] prev = new int[n + 1];
    int[] curr = new int[n + 1];
    for (int j = 0; j <= n; j++) {
      prev[j] = j;
    }
    for (int i = 1; i <= m; i++) {
      curr[0] = i;
      for (int j = 1; j <= n; j++) {
        if (a.charAt(i - 1) == b.charAt(j - 1)) {
          curr[j] = prev[j - 1];
        } else {
          curr[j] = 1 + Math.min(prev[j - 1], Math.min(prev[j], curr[j - 1]));
        }
      }
      int[] tmp = prev;
      prev = curr;
      curr = tmp;
    }
    return prev[n];
  }

  private static Optional<String> terminalContainerName(List<JacksonException.Reference> path) {
    if (path.isEmpty()) {
      return Optional.empty();
    }
    return Optional.ofNullable(path.getLast().getPropertyName());
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
    return "JSON value at '" + renderPath(path) + "' must be an integer value";
  }

  private static String renderPath(List<JacksonException.Reference> path) {
    StringBuilder rendered = new StringBuilder();
    for (JacksonException.Reference reference : path) {
      String name = reference.getPropertyName();
      if (name != null) {
        if (!rendered.isEmpty()) {
          rendered.append('.');
        }
        rendered.append(name);
      } else {
        rendered.append('[').append(reference.getIndex()).append(']');
      }
    }
    return rendered.toString();
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
        jsonPath(exception),
        jsonLine(exception.getLocation()),
        jsonColumn(exception.getLocation()));
  }

  private static Optional<String> jsonPath(JacksonException exception) {
    String rendered = renderPath(exception.getPath());
    return rendered.isEmpty() ? Optional.empty() : Optional.of(rendered);
  }

  private record PayloadMetadata(
      Optional<String> jsonPath, Optional<Integer> jsonLine, Optional<Integer> jsonColumn) {
    private PayloadMetadata {
      Objects.requireNonNull(jsonPath, "jsonPath must not be null");
      Objects.requireNonNull(jsonLine, "jsonLine must not be null");
      Objects.requireNonNull(jsonColumn, "jsonColumn must not be null");
    }
  }
}
