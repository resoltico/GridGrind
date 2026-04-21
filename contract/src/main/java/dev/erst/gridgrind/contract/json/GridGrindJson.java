package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.catalog.GoalPlanReport;
import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import dev.erst.gridgrind.contract.catalog.TaskCatalog;
import dev.erst.gridgrind.contract.catalog.TaskEntry;
import dev.erst.gridgrind.contract.catalog.TaskPlanTemplate;
import dev.erst.gridgrind.contract.catalog.TypeEntry;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import tools.jackson.core.JacksonException;
import tools.jackson.core.StreamReadConstraints;
import tools.jackson.core.StreamReadFeature;
import tools.jackson.core.StreamWriteFeature;
import tools.jackson.core.TokenStreamLocation;
import tools.jackson.core.exc.StreamConstraintsException;
import tools.jackson.core.exc.StreamReadException;
import tools.jackson.core.json.JsonFactory;
import tools.jackson.core.json.JsonFactoryBuilder;
import tools.jackson.databind.DatabindException;
import tools.jackson.databind.DeserializationFeature;
import tools.jackson.databind.SerializationFeature;
import tools.jackson.databind.cfg.CoercionAction;
import tools.jackson.databind.cfg.CoercionInputShape;
import tools.jackson.databind.json.JsonMapper;
import tools.jackson.databind.type.LogicalType;

/** Shared JSON codec for the GridGrind protocol. */
@SuppressWarnings("PMD.ExcessiveImports")
public final class GridGrindJson {
  private static final Pattern MISSING_REQUIRED_FIELD_PATTERN =
      Pattern.compile("missing required creator property '([^']+)'", Pattern.CASE_INSENSITIVE);
  private static final Pattern MISSING_TYPE_ID_FIELD_PATTERN =
      Pattern.compile("missing type id property '([^']+)'", Pattern.CASE_INSENSITIVE);
  private static final Pattern NULL_FIELD_PROBLEM_PATTERN =
      Pattern.compile("problem: ([A-Za-z0-9]+) must not be null", Pattern.CASE_INSENSITIVE);
  private static final JsonMapper JSON_MAPPER = buildMapper(unlimitedJsonFactory());
  private static final JsonMapper REQUEST_JSON_MAPPER = buildMapper(requestJsonFactory());

  private GridGrindJson() {}

  /** Reads a request from an input stream without closing the caller-owned stream. */
  public static WorkbookPlan readRequest(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    try {
      return REQUEST_JSON_MAPPER.readValue(inputStream, WorkbookPlan.class);
    } catch (JacksonException exception) {
      throw invalidRequestPayload(exception);
    }
  }

  /** Reads a request from a byte array. */
  public static WorkbookPlan readRequest(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    requireSupportedRequestLength(bytes.length);
    try {
      return REQUEST_JSON_MAPPER.readValue(bytes, WorkbookPlan.class);
    } catch (JacksonException exception) {
      throw invalidRequestPayload(exception);
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
  public static Catalog readProtocolCatalog(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    try {
      return JSON_MAPPER.readValue(inputStream, Catalog.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a protocol catalog from a byte array. */
  public static Catalog readProtocolCatalog(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    try {
      return JSON_MAPPER.readValue(bytes, Catalog.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a task catalog from an input stream without closing the caller-owned stream. */
  public static TaskCatalog readTaskCatalog(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    try {
      return JSON_MAPPER.readValue(inputStream, TaskCatalog.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a task catalog from a byte array. */
  public static TaskCatalog readTaskCatalog(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    try {
      return JSON_MAPPER.readValue(bytes, TaskCatalog.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a task plan template from an input stream without closing the caller-owned stream. */
  public static TaskPlanTemplate readTaskPlanTemplate(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    try {
      return JSON_MAPPER.readValue(inputStream, TaskPlanTemplate.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a task plan template from a byte array. */
  public static TaskPlanTemplate readTaskPlanTemplate(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    try {
      return JSON_MAPPER.readValue(bytes, TaskPlanTemplate.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a request doctor report from an input stream without closing the caller-owned stream. */
  public static RequestDoctorReport readRequestDoctorReport(InputStream inputStream)
      throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    try {
      return JSON_MAPPER.readValue(inputStream, RequestDoctorReport.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a request doctor report from a byte array. */
  public static RequestDoctorReport readRequestDoctorReport(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    try {
      return JSON_MAPPER.readValue(bytes, RequestDoctorReport.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a goal plan report from an input stream without closing the caller-owned stream. */
  public static GoalPlanReport readGoalPlanReport(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    try {
      return JSON_MAPPER.readValue(inputStream, GoalPlanReport.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Reads a goal plan report from a byte array. */
  public static GoalPlanReport readGoalPlanReport(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    try {
      return JSON_MAPPER.readValue(bytes, GoalPlanReport.class);
    } catch (JacksonException exception) {
      throw invalidPayload(exception);
    }
  }

  /** Serializes a request to bytes. */
  public static byte[] writeRequestBytes(WorkbookPlan request) throws IOException {
    Objects.requireNonNull(request, "request must not be null");
    return writeBytes(request);
  }

  /** Serializes a response to bytes. */
  public static byte[] writeResponseBytes(GridGrindResponse response) throws IOException {
    Objects.requireNonNull(response, "response must not be null");
    return writeBytes(response);
  }

  /** Serializes a protocol catalog to bytes. */
  public static byte[] writeProtocolCatalogBytes(Catalog catalog) throws IOException {
    Objects.requireNonNull(catalog, "catalog must not be null");
    return writeBytes(catalog);
  }

  /** Serializes a task catalog to bytes. */
  public static byte[] writeTaskCatalogBytes(TaskCatalog catalog) throws IOException {
    Objects.requireNonNull(catalog, "catalog must not be null");
    return writeBytes(catalog);
  }

  /** Serializes a task plan template to bytes. */
  public static byte[] writeTaskPlanTemplateBytes(TaskPlanTemplate template) throws IOException {
    Objects.requireNonNull(template, "template must not be null");
    return writeBytes(template);
  }

  /** Serializes a request doctor report to bytes. */
  public static byte[] writeRequestDoctorReportBytes(RequestDoctorReport report)
      throws IOException {
    Objects.requireNonNull(report, "report must not be null");
    return writeBytes(report);
  }

  /** Serializes a goal plan report to bytes. */
  public static byte[] writeGoalPlanReportBytes(GoalPlanReport report) throws IOException {
    Objects.requireNonNull(report, "report must not be null");
    return writeBytes(report);
  }

  /** Writes a request to an output stream without closing the caller-owned stream. */
  public static void writeRequest(OutputStream outputStream, WorkbookPlan request)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(request, "request must not be null");
    JSON_MAPPER.writeValue(outputStream, request);
  }

  /** Writes a response to an output stream without closing the caller-owned stream. */
  public static void writeResponse(OutputStream outputStream, GridGrindResponse response)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(response, "response must not be null");
    JSON_MAPPER.writeValue(outputStream, response);
  }

  /** Writes a protocol catalog to an output stream without closing the caller-owned stream. */
  public static void writeProtocolCatalog(OutputStream outputStream, Catalog catalog)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(catalog, "catalog must not be null");
    JSON_MAPPER.writeValue(outputStream, catalog);
  }

  /** Writes a task catalog to an output stream without closing the caller-owned stream. */
  public static void writeTaskCatalog(OutputStream outputStream, TaskCatalog catalog)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(catalog, "catalog must not be null");
    JSON_MAPPER.writeValue(outputStream, catalog);
  }

  /** Writes a task plan template to an output stream without closing the caller-owned stream. */
  public static void writeTaskPlanTemplate(OutputStream outputStream, TaskPlanTemplate template)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(template, "template must not be null");
    JSON_MAPPER.writeValue(outputStream, template);
  }

  /** Writes a request doctor report to an output stream without closing the caller-owned stream. */
  public static void writeRequestDoctorReport(OutputStream outputStream, RequestDoctorReport report)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(report, "report must not be null");
    JSON_MAPPER.writeValue(outputStream, report);
  }

  /** Writes a goal plan report to an output stream without closing the caller-owned stream. */
  public static void writeGoalPlanReport(OutputStream outputStream, GoalPlanReport report)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(report, "report must not be null");
    JSON_MAPPER.writeValue(outputStream, report);
  }

  /**
   * Writes a single catalog type entry to an output stream without closing the caller-owned stream.
   */
  public static void writeTypeEntry(OutputStream outputStream, TypeEntry entry) throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(entry, "entry must not be null");
    JSON_MAPPER.writeValue(outputStream, entry);
  }

  /** Writes a single task entry to an output stream without closing the caller-owned stream. */
  public static void writeTaskEntry(OutputStream outputStream, TaskEntry entry) throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(entry, "entry must not be null");
    JSON_MAPPER.writeValue(outputStream, entry);
  }

  /** Returns the maximum accepted JSON request document length in bytes. */
  public static long maxRequestDocumentBytes() {
    return GridGrindContractText.requestDocumentLimitBytes();
  }

  /** Rejects one request payload length that exceeds the documented transport limit. */
  public static void requireSupportedRequestLength(long lengthBytes) {
    if (lengthBytes > maxRequestDocumentBytes()) {
      throw requestTooLarge(null);
    }
  }

  static IllegalArgumentException invalidPayload(JacksonException exception) {
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

  private static IllegalArgumentException invalidRequestPayload(JacksonException exception) {
    if (exception instanceof StreamConstraintsException) {
      return requestTooLarge(exception);
    }
    return invalidPayload(exception);
  }

  private static Throwable validationCause(Throwable throwable) {
    Throwable current = throwable;
    while (current != null) {
      if (current instanceof IllegalArgumentException
          || current instanceof java.time.DateTimeException) {
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
    if (throwable
        instanceof tools.jackson.databind.exc.InvalidTypeIdException invalidTypeIdException) {
      String typeId = invalidTypeIdException.getTypeId();
      return "Unknown type value '" + typeId + "'";
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

  /**
   * Returns the public wording for mismatched JSON token shapes.
   *
   * <p>When a required primitive field is absent or supplied as JSON {@code null}, Jackson maps it
   * as a null-into-primitive coercion failure rather than a missing-creator-property failure. The
   * raw message leaks internal Jackson configuration advice; extract the field name from the
   * exception path and surface it as a standard missing-required-field message instead.
   */
  @SuppressWarnings("StringConcatToTextBlock")
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

  /**
   * Removes Jackson-specific noise from one parser message.
   *
   * <p>Strips four categories of Jackson noise: - Source-location suffix: {@code (start marker at
   * [Source:...])} - Subtype description: {@code as a subtype of `X`} - POJO property reference:
   * {@code (for POJO property 'X')} - Configuration advice: {@code (set X.Y to 'Z' to allow)} — the
   * universal pattern Jackson uses to suggest enabling or disabling deserialization features
   *
   * <p>The fallback {@code return cleaned} in {@link #productOwnedJacksonMessage} is only safe
   * because this method guarantees the output is noise-free. Any new Jackson noise pattern
   * discovered in a real error message must be added here as a stripping rule.
   */
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

  /* Returns the 1-based line number from the location, or null when unavailable. */
  static Integer jsonLine(TokenStreamLocation location) {
    if (location == null) {
      return null;
    }
    int line = location.getLineNr();
    return line > 0 ? line : null;
  }

  /* Returns the 1-based column number from the location, or null when unavailable. */
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

  private static JsonMapper buildMapper(JsonFactory jsonFactory) {
    return JsonMapper.builder(jsonFactory)
        .disable(StreamReadFeature.AUTO_CLOSE_SOURCE)
        .disable(StreamWriteFeature.AUTO_CLOSE_TARGET)
        .enable(SerializationFeature.INDENT_OUTPUT)
        .enable(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES)
        .withCoercionConfig(
            LogicalType.Integer,
            config -> config.setCoercion(CoercionInputShape.Float, CoercionAction.Fail))
        .build();
  }

  private static JsonFactory unlimitedJsonFactory() {
    return new JsonFactoryBuilder().build();
  }

  private static JsonFactory requestJsonFactory() {
    return new JsonFactoryBuilder()
        .streamReadConstraints(
            StreamReadConstraints.builder().maxDocumentLength(maxRequestDocumentBytes()).build())
        .build();
  }

  private static InvalidRequestException requestTooLarge(Throwable cause) {
    return new InvalidRequestException(
        GridGrindContractText.requestDocumentTooLargeMessage(), null, null, null, cause);
  }

  private record PayloadMetadata(String jsonPath, Integer jsonLine, Integer jsonColumn) {}
}
