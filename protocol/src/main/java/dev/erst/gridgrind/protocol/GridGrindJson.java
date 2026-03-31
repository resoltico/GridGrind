package dev.erst.gridgrind.protocol;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.time.DateTimeException;
import java.util.Objects;
import java.util.regex.Pattern;
import tools.jackson.core.JacksonException;
import tools.jackson.core.StreamReadFeature;
import tools.jackson.core.StreamWriteFeature;
import tools.jackson.core.TokenStreamLocation;
import tools.jackson.core.exc.StreamReadException;
import tools.jackson.databind.DatabindException;
import tools.jackson.databind.DeserializationFeature;
import tools.jackson.databind.SerializationFeature;
import tools.jackson.databind.exc.InvalidTypeIdException;
import tools.jackson.databind.exc.MismatchedInputException;
import tools.jackson.databind.exc.UnrecognizedPropertyException;
import tools.jackson.databind.json.JsonMapper;

/** Shared JSON codec for the GridGrind protocol. */
public final class GridGrindJson {
  private static final Pattern MISSING_REQUIRED_FIELD_PATTERN =
      Pattern.compile("missing required creator property '([^']+)'", Pattern.CASE_INSENSITIVE);
  private static final Pattern MISSING_TYPE_ID_FIELD_PATTERN =
      Pattern.compile("missing type id property '([^']+)'", Pattern.CASE_INSENSITIVE);
  private static final Pattern NULL_FIELD_PROBLEM_PATTERN =
      Pattern.compile("problem: ([A-Za-z0-9]+) must not be null", Pattern.CASE_INSENSITIVE);
  private static final JsonMapper JSON_MAPPER =
      JsonMapper.builder()
          .disable(StreamReadFeature.AUTO_CLOSE_SOURCE)
          .disable(StreamWriteFeature.AUTO_CLOSE_TARGET)
          .enable(SerializationFeature.INDENT_OUTPUT)
          .enable(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES)
          .build();

  private GridGrindJson() {}

  /** Reads a request from an input stream without closing the caller-owned stream. */
  public static GridGrindRequest readRequest(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    try {
      return JSON_MAPPER.readValue(inputStream, GridGrindRequest.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a request from a byte array. */
  public static GridGrindRequest readRequest(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    try {
      return JSON_MAPPER.readValue(bytes, GridGrindRequest.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a response from an input stream without closing the caller-owned stream. */
  public static GridGrindResponse readResponse(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    try {
      return JSON_MAPPER.readValue(inputStream, GridGrindResponse.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a response from a byte array. */
  public static GridGrindResponse readResponse(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    try {
      return JSON_MAPPER.readValue(bytes, GridGrindResponse.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a protocol catalog from an input stream without closing the caller-owned stream. */
  public static GridGrindProtocolCatalog.Catalog readProtocolCatalog(InputStream inputStream)
      throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    try {
      return JSON_MAPPER.readValue(inputStream, GridGrindProtocolCatalog.Catalog.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a protocol catalog from a byte array. */
  public static GridGrindProtocolCatalog.Catalog readProtocolCatalog(byte[] bytes)
      throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    try {
      return JSON_MAPPER.readValue(bytes, GridGrindProtocolCatalog.Catalog.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Serializes a request to bytes. */
  public static byte[] writeRequestBytes(GridGrindRequest request) throws IOException {
    Objects.requireNonNull(request, "request must not be null");
    return writeBytes(request);
  }

  /** Serializes a response to bytes. */
  public static byte[] writeResponseBytes(GridGrindResponse response) throws IOException {
    Objects.requireNonNull(response, "response must not be null");
    return writeBytes(response);
  }

  /** Serializes a protocol catalog to bytes. */
  public static byte[] writeProtocolCatalogBytes(GridGrindProtocolCatalog.Catalog catalog)
      throws IOException {
    Objects.requireNonNull(catalog, "catalog must not be null");
    return writeBytes(catalog);
  }

  /** Writes a request to an output stream without closing the caller-owned stream. */
  public static void writeRequest(OutputStream outputStream, GridGrindRequest request)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    outputStream.write(writeRequestBytes(request));
  }

  /** Writes a response to an output stream without closing the caller-owned stream. */
  public static void writeResponse(OutputStream outputStream, GridGrindResponse response)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    outputStream.write(writeResponseBytes(response));
  }

  /** Writes a protocol catalog to an output stream without closing the caller-owned stream. */
  public static void writeProtocolCatalog(
      OutputStream outputStream, GridGrindProtocolCatalog.Catalog catalog) throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    outputStream.write(writeProtocolCatalogBytes(catalog));
  }

  /**
   * Writes a single catalog type entry to an output stream without closing the caller-owned stream.
   */
  public static void writeTypeEntry(
      OutputStream outputStream, GridGrindProtocolCatalog.TypeEntry entry) throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(entry, "entry must not be null");
    outputStream.write(JSON_MAPPER.writeValueAsBytes(entry));
  }

  private static IllegalArgumentException invalidPayload(JacksonException exception) {
    PayloadMetadata metadata = payloadMetadata(exception);
    Throwable validationCause = validationCause(exception);
    if (exception instanceof StreamReadException) {
      return new InvalidJsonException(
          message(exception),
          metadata.jsonPath(),
          metadata.jsonLine(),
          metadata.jsonColumn(),
          exception);
    }
    if (validationCause != null) {
      return new InvalidRequestException(
          message(validationCause),
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

  private static Throwable validationCause(Throwable throwable) {
    Throwable current = throwable;
    while (current != null) {
      if (current instanceof IllegalArgumentException || current instanceof DateTimeException) {
        return current;
      }
      current = current.getCause();
    }
    return null;
  }

  /**
   * Returns the public invalid-request-shape message for one parsing failure.
   *
   * <p>Package-private so tests can assert the exact public wording of otherwise hard-to-trigger
   * parser branches.
   */
  static String message(Throwable throwable) {
    if (throwable instanceof InvalidTypeIdException invalidTypeIdException) {
      return "Unknown type value '" + invalidTypeIdException.getTypeId() + "'";
    }
    if (throwable instanceof UnrecognizedPropertyException unrecognizedPropertyException) {
      return "Unknown field '" + unrecognizedPropertyException.getPropertyName() + "'";
    }
    if (throwable instanceof MismatchedInputException mismatchedInputException) {
      return mismatchedInputMessage(mismatchedInputException);
    }
    String message =
        throwable instanceof JacksonException jacksonException
            ? jacksonException.getOriginalMessage()
            : throwable.getMessage();
    return productOwnedJacksonMessage(cleanJacksonMessage(message));
  }

  /** Returns the public wording for mismatched JSON token shapes. */
  static String mismatchedInputMessage(MismatchedInputException exception) {
    return productOwnedJacksonMessage(cleanJacksonMessage(exception.getOriginalMessage()));
  }

  private static String productOwnedJacksonMessage(String cleaned) {
    java.util.regex.Matcher missingRequiredField = MISSING_REQUIRED_FIELD_PATTERN.matcher(cleaned);
    if (missingRequiredField.find()) {
      return "Missing required field '" + missingRequiredField.group(1) + "'";
    }
    java.util.regex.Matcher missingTypeIdField = MISSING_TYPE_ID_FIELD_PATTERN.matcher(cleaned);
    if (missingTypeIdField.find()) {
      return "Missing required field '" + missingTypeIdField.group(1) + "'";
    }
    java.util.regex.Matcher nullFieldProblem = NULL_FIELD_PROBLEM_PATTERN.matcher(cleaned);
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

  /** Removes Jackson-specific suffix noise from one parser message. */
  static String cleanJacksonMessage(String message) {
    if (message == null || message.isBlank()) {
      return "Invalid JSON payload";
    }
    int startMarkerIndex = message.indexOf(" (start marker at [Source:");
    String trimmed = startMarkerIndex >= 0 ? message.substring(0, startMarkerIndex) : message;
    return trimmed
        .replaceAll(" as a subtype of `[^`]+`", "")
        .replaceAll(" \\(for POJO property '[^']+'\\)", "");
  }

  private static PayloadMetadata payloadMetadata(JacksonException exception) {
    return new PayloadMetadata(
        jsonPath(exception),
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

  /** Returns the 1-based line number from the location, or null when unavailable. */
  static Integer jsonLine(TokenStreamLocation location) {
    if (location == null) {
      return null;
    }
    int line = location.getLineNr();
    return line > 0 ? line : null;
  }

  /** Returns the 1-based column number from the location, or null when unavailable. */
  static Integer jsonColumn(TokenStreamLocation location) {
    if (location == null) {
      return null;
    }
    int column = location.getColumnNr();
    return column > 0 ? column : null;
  }

  private static byte[] writeBytes(Object value) throws IOException {
    return JSON_MAPPER.writeValueAsBytes(value);
  }

  private record PayloadMetadata(String jsonPath, Integer jsonLine, Integer jsonColumn) {}
}
