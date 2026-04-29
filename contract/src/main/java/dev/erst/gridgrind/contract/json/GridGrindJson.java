package dev.erst.gridgrind.contract.json;

import com.fasterxml.jackson.annotation.JsonInclude;
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
import java.util.Optional;
import java.util.function.Function;
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
import tools.jackson.databind.JsonNode;
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
      Pattern.compile("problem: ([A-Za-z0-9.\\[\\]_]+) must not be null", Pattern.CASE_INSENSITIVE);
  private static final JsonMapper JSON_MAPPER = buildMapper(unlimitedJsonFactory());
  private static final JsonMapper WIRE_JSON_MAPPER = buildMapper(unlimitedJsonFactory(), true);
  private static final JsonMapper DISCOVERY_JSON_MAPPER = buildMapper(unlimitedJsonFactory(), true);
  private static final JsonMapper REQUEST_JSON_MAPPER = buildMapper(requestJsonFactory());

  private GridGrindJson() {}

  /** Reads a request from an input stream without closing the caller-owned stream. */
  public static WorkbookPlan readRequest(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return readValue(
        inputStream, REQUEST_JSON_MAPPER, WorkbookPlan.class, GridGrindJson::invalidRequestPayload);
  }

  /** Reads a request from a byte array. */
  public static WorkbookPlan readRequest(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    requireSupportedRequestLength(bytes.length);
    return readValue(
        bytes, REQUEST_JSON_MAPPER, WorkbookPlan.class, GridGrindJson::invalidRequestPayload);
  }

  /** Reads a response from an input stream without closing the caller-owned stream. */
  public static GridGrindResponse readResponse(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return readValue(
        inputStream, JSON_MAPPER, GridGrindResponse.class, GridGrindJson::invalidPayload);
  }

  /** Reads a response from a byte array. */
  public static GridGrindResponse readResponse(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return readValue(bytes, JSON_MAPPER, GridGrindResponse.class, GridGrindJson::invalidPayload);
  }

  /** Reads a protocol catalog from an input stream without closing the caller-owned stream. */
  public static Catalog readProtocolCatalog(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return readValue(inputStream, JSON_MAPPER, Catalog.class, GridGrindJson::invalidPayload);
  }

  /** Reads a protocol catalog from a byte array. */
  public static Catalog readProtocolCatalog(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return readValue(bytes, JSON_MAPPER, Catalog.class, GridGrindJson::invalidPayload);
  }

  /** Reads a task catalog from an input stream without closing the caller-owned stream. */
  public static TaskCatalog readTaskCatalog(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return readValue(inputStream, JSON_MAPPER, TaskCatalog.class, GridGrindJson::invalidPayload);
  }

  /** Reads a task catalog from a byte array. */
  public static TaskCatalog readTaskCatalog(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return readValue(bytes, JSON_MAPPER, TaskCatalog.class, GridGrindJson::invalidPayload);
  }

  /** Reads a task plan template from an input stream without closing the caller-owned stream. */
  public static TaskPlanTemplate readTaskPlanTemplate(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return readValue(
        inputStream, JSON_MAPPER, TaskPlanTemplate.class, GridGrindJson::invalidPayload);
  }

  /** Reads a task plan template from a byte array. */
  public static TaskPlanTemplate readTaskPlanTemplate(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return readValue(bytes, JSON_MAPPER, TaskPlanTemplate.class, GridGrindJson::invalidPayload);
  }

  /** Reads a request doctor report from an input stream without closing the caller-owned stream. */
  public static RequestDoctorReport readRequestDoctorReport(InputStream inputStream)
      throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return readValue(
        inputStream, JSON_MAPPER, RequestDoctorReport.class, GridGrindJson::invalidPayload);
  }

  /** Reads a request doctor report from a byte array. */
  public static RequestDoctorReport readRequestDoctorReport(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return readValue(bytes, JSON_MAPPER, RequestDoctorReport.class, GridGrindJson::invalidPayload);
  }

  /** Reads a goal plan report from an input stream without closing the caller-owned stream. */
  public static GoalPlanReport readGoalPlanReport(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return readValue(inputStream, JSON_MAPPER, GoalPlanReport.class, GridGrindJson::invalidPayload);
  }

  /** Reads a goal plan report from a byte array. */
  public static GoalPlanReport readGoalPlanReport(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return readValue(bytes, JSON_MAPPER, GoalPlanReport.class, GridGrindJson::invalidPayload);
  }

  /** Serializes a request to bytes. */
  public static byte[] writeRequestBytes(WorkbookPlan request) throws IOException {
    Objects.requireNonNull(request, "request must not be null");
    return writeWireBytes(request);
  }

  /** Serializes a response to bytes. */
  public static byte[] writeResponseBytes(GridGrindResponse response) throws IOException {
    Objects.requireNonNull(response, "response must not be null");
    return writeWireBytes(response);
  }

  /** Serializes a protocol catalog to bytes. */
  public static byte[] writeProtocolCatalogBytes(Catalog catalog) throws IOException {
    Objects.requireNonNull(catalog, "catalog must not be null");
    return writeDiscoveryBytes(catalog);
  }

  /** Serializes a task catalog to bytes. */
  public static byte[] writeTaskCatalogBytes(TaskCatalog catalog) throws IOException {
    Objects.requireNonNull(catalog, "catalog must not be null");
    return writeDiscoveryBytes(catalog);
  }

  /** Serializes a task plan template to bytes. */
  public static byte[] writeTaskPlanTemplateBytes(TaskPlanTemplate template) throws IOException {
    Objects.requireNonNull(template, "template must not be null");
    return writeDiscoveryBytes(template);
  }

  /** Serializes a request doctor report to bytes. */
  public static byte[] writeRequestDoctorReportBytes(RequestDoctorReport report)
      throws IOException {
    Objects.requireNonNull(report, "report must not be null");
    return writeDiscoveryBytes(report);
  }

  /** Serializes a goal plan report to bytes. */
  public static byte[] writeGoalPlanReportBytes(GoalPlanReport report) throws IOException {
    Objects.requireNonNull(report, "report must not be null");
    return writeDiscoveryBytes(report);
  }

  /** Writes a request to an output stream without closing the caller-owned stream. */
  public static void writeRequest(OutputStream outputStream, WorkbookPlan request)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(request, "request must not be null");
    WIRE_JSON_MAPPER.writeValue(outputStream, request);
  }

  /** Writes a response to an output stream without closing the caller-owned stream. */
  public static void writeResponse(OutputStream outputStream, GridGrindResponse response)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(response, "response must not be null");
    WIRE_JSON_MAPPER.writeValue(outputStream, response);
  }

  /** Writes a protocol catalog to an output stream without closing the caller-owned stream. */
  public static void writeProtocolCatalog(OutputStream outputStream, Catalog catalog)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(catalog, "catalog must not be null");
    DISCOVERY_JSON_MAPPER.writeValue(outputStream, catalog);
  }

  /** Writes a task catalog to an output stream without closing the caller-owned stream. */
  public static void writeTaskCatalog(OutputStream outputStream, TaskCatalog catalog)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(catalog, "catalog must not be null");
    DISCOVERY_JSON_MAPPER.writeValue(outputStream, catalog);
  }

  /** Writes a task plan template to an output stream without closing the caller-owned stream. */
  public static void writeTaskPlanTemplate(OutputStream outputStream, TaskPlanTemplate template)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(template, "template must not be null");
    DISCOVERY_JSON_MAPPER.writeValue(outputStream, template);
  }

  /** Writes a request doctor report to an output stream without closing the caller-owned stream. */
  public static void writeRequestDoctorReport(OutputStream outputStream, RequestDoctorReport report)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(report, "report must not be null");
    DISCOVERY_JSON_MAPPER.writeValue(outputStream, report);
  }

  /** Writes a goal plan report to an output stream without closing the caller-owned stream. */
  public static void writeGoalPlanReport(OutputStream outputStream, GoalPlanReport report)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(report, "report must not be null");
    DISCOVERY_JSON_MAPPER.writeValue(outputStream, report);
  }

  /**
   * Writes a single catalog type entry to an output stream without closing the caller-owned stream.
   */
  public static void writeTypeEntry(OutputStream outputStream, TypeEntry entry) throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(entry, "entry must not be null");
    DISCOVERY_JSON_MAPPER.writeValue(outputStream, entry);
  }

  /** Writes one protocol-catalog lookup value to an output stream without closing it. */
  public static void writeCatalogLookupValue(OutputStream outputStream, Object value)
      throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(value, "value must not be null");
    DISCOVERY_JSON_MAPPER.writeValue(outputStream, value);
  }

  /** Writes a single task entry to an output stream without closing the caller-owned stream. */
  public static void writeTaskEntry(OutputStream outputStream, TaskEntry entry) throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(entry, "entry must not be null");
    DISCOVERY_JSON_MAPPER.writeValue(outputStream, entry);
  }

  /** Returns the maximum accepted JSON request document length in bytes. */
  public static long maxRequestDocumentBytes() {
    return GridGrindContractText.requestDocumentLimitBytes();
  }

  /** Rejects one request payload length that exceeds the documented transport limit. */
  public static void requireSupportedRequestLength(long lengthBytes) {
    // LIM-021
    if (lengthBytes > maxRequestDocumentBytes()) {
      throw requestTooLarge(null);
    }
  }

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

  private static IllegalArgumentException invalidRequestPayload(JacksonException exception) {
    if (exception instanceof StreamConstraintsException) {
      return requestTooLarge(exception);
    }
    return invalidPayload(exception);
  }

  private static Optional<Throwable> validationCause(Throwable throwable) {
    Throwable current = throwable;
    while (current != null) {
      if (current instanceof IllegalArgumentException
          || current instanceof java.time.DateTimeException) {
        return Optional.of(current);
      }
      current = current.getCause();
    }
    return Optional.empty();
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
                + " EXPECT_TABLE_PRESENT, EXPECT_PIVOT_TABLE_PRESENT, or"
                + " EXPECT_CHART_PRESENT";
        case "EXPECT_ABSENT" ->
            defaultMessage
                + "; use one explicit family assertion such as EXPECT_NAMED_RANGE_ABSENT,"
                + " EXPECT_TABLE_ABSENT, EXPECT_PIVOT_TABLE_ABSENT, or"
                + " EXPECT_CHART_ABSENT";
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
        jsonLine(exception.getLocation()).orElse(null),
        jsonColumn(exception.getLocation()).orElse(null));
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
  static Optional<Integer> jsonLine(TokenStreamLocation location) {
    if (location == null) {
      return Optional.empty();
    }
    int line = location.getLineNr();
    return line > 0 ? Optional.of(line) : Optional.empty();
  }

  /* Returns the 1-based column number from the location, or null when unavailable. */
  static Optional<Integer> jsonColumn(TokenStreamLocation location) {
    if (location == null) {
      return Optional.empty();
    }
    int column = location.getColumnNr();
    return column > 0 ? Optional.of(column) : Optional.empty();
  }

  private static byte[] writeWireBytes(Object value) throws IOException {
    return WIRE_JSON_MAPPER.writeValueAsBytes(value);
  }

  private static byte[] writeDiscoveryBytes(Object value) throws IOException {
    return DISCOVERY_JSON_MAPPER.writeValueAsBytes(value);
  }

  private static <T> T readValue(
      InputStream inputStream,
      JsonMapper mapper,
      Class<T> targetType,
      Function<JacksonException, IllegalArgumentException> failureMapper)
      throws IOException {
    return decodeValue(
        readTree(inputStream, mapper, failureMapper), mapper, targetType, failureMapper);
  }

  private static <T> T readValue(
      byte[] bytes,
      JsonMapper mapper,
      Class<T> targetType,
      Function<JacksonException, IllegalArgumentException> failureMapper)
      throws IOException {
    return decodeValue(readTree(bytes, mapper, failureMapper), mapper, targetType, failureMapper);
  }

  private static JsonNode readTree(
      InputStream inputStream,
      JsonMapper mapper,
      Function<JacksonException, IllegalArgumentException> failureMapper)
      throws IOException {
    try {
      return requirePresentDocument(mapper.readTree(inputStream));
    } catch (JacksonException exception) {
      throw failureMapper.apply(exception);
    }
  }

  private static JsonNode readTree(
      byte[] bytes,
      JsonMapper mapper,
      Function<JacksonException, IllegalArgumentException> failureMapper)
      throws IOException {
    try {
      return requirePresentDocument(mapper.readTree(bytes));
    } catch (JacksonException exception) {
      throw failureMapper.apply(exception);
    }
  }

  private static <T> T decodeValue(
      JsonNode node,
      JsonMapper mapper,
      Class<T> targetType,
      Function<JacksonException, IllegalArgumentException> failureMapper)
      throws IOException {
    requireNonNullRoot(node);
    try {
      rejectExplicitNullMembers(node, "");
      return mapper.treeToValue(node, targetType);
    } catch (JacksonException exception) {
      throw failureMapper.apply(exception);
    } catch (IllegalArgumentException exception) {
      throw new InvalidRequestException(message(exception), null, null, null, exception);
    }
  }

  private static void rejectExplicitNullMembers(JsonNode node, String path) {
    if (node.isObject()) {
      for (var entry : node.properties()) {
        String childPath = path.isEmpty() ? entry.getKey() : path + "." + entry.getKey();
        if (entry.getValue().isNull()) {
          throw new IllegalArgumentException("problem: " + childPath + " must not be null");
        }
        rejectExplicitNullMembers(entry.getValue(), childPath);
      }
      return;
    }
    if (node.isArray()) {
      for (int index = 0; index < node.size(); index++) {
        JsonNode child = node.get(index);
        String childPath = path + "[" + index + "]";
        if (child.isNull()) {
          throw new IllegalArgumentException("problem: " + childPath + " must not be null");
        }
        rejectExplicitNullMembers(child, childPath);
      }
    }
  }

  private static void requireNonNullRoot(JsonNode node) {
    if (node.isNull()) {
      throw new InvalidRequestException("problem: <root> must not be null", null, null, null, null);
    }
  }

  private static JsonNode requirePresentDocument(JsonNode node) {
    return Optional.ofNullable(node)
        .filter(candidate -> !candidate.isMissingNode())
        .orElseThrow(GridGrindJson::invalidJsonPayloadException);
  }

  private static InvalidJsonException invalidJsonPayloadException() {
    return new InvalidJsonException("Invalid JSON payload", null, null, null, null);
  }

  private static JsonMapper buildMapper(JsonFactory jsonFactory) {
    return buildMapper(jsonFactory, false);
  }

  private static JsonMapper buildMapper(JsonFactory jsonFactory, boolean omitNullValues) {
    JsonMapper.Builder builder =
        JsonMapper.builder(jsonFactory)
            .disable(StreamReadFeature.AUTO_CLOSE_SOURCE)
            .disable(StreamWriteFeature.AUTO_CLOSE_TARGET)
            .enable(SerializationFeature.INDENT_OUTPUT)
            .enable(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES)
            .withCoercionConfig(
                LogicalType.Integer,
                config -> config.setCoercion(CoercionInputShape.Float, CoercionAction.Fail));
    if (omitNullValues) {
      builder.changeDefaultPropertyInclusion(
          value -> value.withValueInclusion(JsonInclude.Include.NON_ABSENT));
    }
    return builder.build();
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
