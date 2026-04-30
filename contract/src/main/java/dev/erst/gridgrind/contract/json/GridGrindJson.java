package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.catalog.GoalPlanReport;
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
import java.util.Objects;
import tools.jackson.core.JacksonException;
import tools.jackson.core.TokenStreamLocation;
import tools.jackson.databind.JsonNode;

/** Shared JSON codec for the GridGrind protocol. */
public final class GridGrindJson {
  private GridGrindJson() {}

  /** Reads a request from an input stream without closing the caller-owned stream. */
  public static WorkbookPlan readRequest(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return readRequestTree(
        GridGrindJsonCodecSupport.readTree(
            inputStream,
            GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER,
            GridGrindJsonMessageSupport::invalidRequestPayload));
  }

  /** Reads a request from a byte array. */
  public static WorkbookPlan readRequest(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    GridGrindJsonMapperSupport.requireSupportedRequestLength(bytes.length);
    return readRequestTree(
        GridGrindJsonCodecSupport.readTree(
            bytes,
            GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER,
            GridGrindJsonMessageSupport::invalidRequestPayload));
  }

  /** Reads a response from an input stream without closing the caller-owned stream. */
  public static GridGrindResponse readResponse(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readValue(
        inputStream,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        GridGrindResponse.class,
        GridGrindJsonMessageSupport::invalidPayload);
  }

  /** Reads a response from a byte array. */
  public static GridGrindResponse readResponse(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.readValue(
        bytes,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        GridGrindResponse.class,
        GridGrindJsonMessageSupport::invalidPayload);
  }

  /** Reads a protocol catalog from an input stream without closing the caller-owned stream. */
  public static Catalog readProtocolCatalog(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readValue(
        inputStream,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        Catalog.class,
        GridGrindJsonMessageSupport::invalidPayload);
  }

  /** Reads a protocol catalog from a byte array. */
  public static Catalog readProtocolCatalog(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.readValue(
        bytes,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        Catalog.class,
        GridGrindJsonMessageSupport::invalidPayload);
  }

  /** Reads a task catalog from an input stream without closing the caller-owned stream. */
  public static TaskCatalog readTaskCatalog(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readValue(
        inputStream,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        TaskCatalog.class,
        GridGrindJsonMessageSupport::invalidPayload);
  }

  /** Reads a task catalog from a byte array. */
  public static TaskCatalog readTaskCatalog(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.readValue(
        bytes,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        TaskCatalog.class,
        GridGrindJsonMessageSupport::invalidPayload);
  }

  /** Reads a task plan template from an input stream without closing the caller-owned stream. */
  public static TaskPlanTemplate readTaskPlanTemplate(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readValue(
        inputStream,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        TaskPlanTemplate.class,
        GridGrindJsonMessageSupport::invalidPayload);
  }

  /** Reads a task plan template from a byte array. */
  public static TaskPlanTemplate readTaskPlanTemplate(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.readValue(
        bytes,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        TaskPlanTemplate.class,
        GridGrindJsonMessageSupport::invalidPayload);
  }

  /** Reads a request doctor report from an input stream without closing the caller-owned stream. */
  public static RequestDoctorReport readRequestDoctorReport(InputStream inputStream)
      throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readValue(
        inputStream,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        RequestDoctorReport.class,
        GridGrindJsonMessageSupport::invalidPayload);
  }

  /** Reads a request doctor report from a byte array. */
  public static RequestDoctorReport readRequestDoctorReport(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.readValue(
        bytes,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        RequestDoctorReport.class,
        GridGrindJsonMessageSupport::invalidPayload);
  }

  /** Reads a goal plan report from an input stream without closing the caller-owned stream. */
  public static GoalPlanReport readGoalPlanReport(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readValue(
        inputStream,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        GoalPlanReport.class,
        GridGrindJsonMessageSupport::invalidPayload);
  }

  /** Reads a goal plan report from a byte array. */
  public static GoalPlanReport readGoalPlanReport(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.readValue(
        bytes,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        GoalPlanReport.class,
        GridGrindJsonMessageSupport::invalidPayload);
  }

  /** Serializes a request to bytes. */
  public static byte[] writeRequestBytes(WorkbookPlan request) throws IOException {
    Objects.requireNonNull(request, "request must not be null");
    return GridGrindJsonCodecSupport.writeBytes(
        GridGrindJsonMapperSupport.WIRE_JSON_MAPPER, request);
  }

  /** Serializes a response to bytes. */
  public static byte[] writeResponseBytes(GridGrindResponse response) throws IOException {
    Objects.requireNonNull(response, "response must not be null");
    return GridGrindJsonCodecSupport.writeBytes(
        GridGrindJsonMapperSupport.WIRE_JSON_MAPPER, response);
  }

  /** Serializes a protocol catalog to bytes. */
  public static byte[] writeProtocolCatalogBytes(Catalog catalog) throws IOException {
    Objects.requireNonNull(catalog, "catalog must not be null");
    return GridGrindJsonCodecSupport.writeBytes(
        GridGrindJsonMapperSupport.WIRE_JSON_MAPPER, catalog);
  }

  /** Serializes a task catalog to bytes. */
  public static byte[] writeTaskCatalogBytes(TaskCatalog catalog) throws IOException {
    Objects.requireNonNull(catalog, "catalog must not be null");
    return GridGrindJsonCodecSupport.writeBytes(
        GridGrindJsonMapperSupport.WIRE_JSON_MAPPER, catalog);
  }

  /** Serializes a task plan template to bytes. */
  public static byte[] writeTaskPlanTemplateBytes(TaskPlanTemplate template) throws IOException {
    Objects.requireNonNull(template, "template must not be null");
    return GridGrindJsonCodecSupport.writeBytes(
        GridGrindJsonMapperSupport.WIRE_JSON_MAPPER, template);
  }

  /** Serializes a request doctor report to bytes. */
  public static byte[] writeRequestDoctorReportBytes(RequestDoctorReport report)
      throws IOException {
    Objects.requireNonNull(report, "report must not be null");
    return GridGrindJsonCodecSupport.writeBytes(
        GridGrindJsonMapperSupport.WIRE_JSON_MAPPER, report);
  }

  /** Serializes a goal plan report to bytes. */
  public static byte[] writeGoalPlanReportBytes(GoalPlanReport report) throws IOException {
    Objects.requireNonNull(report, "report must not be null");
    return GridGrindJsonCodecSupport.writeBytes(
        GridGrindJsonMapperSupport.WIRE_JSON_MAPPER, report);
  }

  /** Writes a request to an output stream without closing the caller-owned stream. */
  public static void writeRequest(OutputStream outputStream, WorkbookPlan request)
      throws IOException {
    writeValue(outputStream, request);
  }

  /** Writes a response to an output stream without closing the caller-owned stream. */
  public static void writeResponse(OutputStream outputStream, GridGrindResponse response)
      throws IOException {
    writeValue(outputStream, response);
  }

  /** Writes a protocol catalog to an output stream without closing the caller-owned stream. */
  public static void writeProtocolCatalog(OutputStream outputStream, Catalog catalog)
      throws IOException {
    writeValue(outputStream, catalog);
  }

  /** Writes a task catalog to an output stream without closing the caller-owned stream. */
  public static void writeTaskCatalog(OutputStream outputStream, TaskCatalog catalog)
      throws IOException {
    writeValue(outputStream, catalog);
  }

  /** Writes a task plan template to an output stream without closing the caller-owned stream. */
  public static void writeTaskPlanTemplate(OutputStream outputStream, TaskPlanTemplate template)
      throws IOException {
    writeValue(outputStream, template);
  }

  /** Writes a request doctor report to an output stream without closing the caller-owned stream. */
  public static void writeRequestDoctorReport(OutputStream outputStream, RequestDoctorReport report)
      throws IOException {
    writeValue(outputStream, report);
  }

  /** Writes a goal plan report to an output stream without closing the caller-owned stream. */
  public static void writeGoalPlanReport(OutputStream outputStream, GoalPlanReport report)
      throws IOException {
    writeValue(outputStream, report);
  }

  /**
   * Writes a single catalog type entry to an output stream without closing the caller-owned stream.
   */
  public static void writeTypeEntry(OutputStream outputStream, TypeEntry entry) throws IOException {
    writeValue(outputStream, entry);
  }

  /** Writes one protocol-catalog lookup value to an output stream without closing it. */
  public static void writeCatalogLookupValue(OutputStream outputStream, Object value)
      throws IOException {
    writeValue(outputStream, value);
  }

  /** Writes a single task entry to an output stream without closing the caller-owned stream. */
  public static void writeTaskEntry(OutputStream outputStream, TaskEntry entry) throws IOException {
    writeValue(outputStream, entry);
  }

  /** Returns the maximum accepted JSON request document length in bytes. */
  public static long maxRequestDocumentBytes() {
    return GridGrindJsonMapperSupport.maxRequestDocumentBytes();
  }

  /** Rejects one request payload length that exceeds the documented transport limit. */
  public static void requireSupportedRequestLength(long lengthBytes) {
    GridGrindJsonMapperSupport.requireSupportedRequestLength(lengthBytes);
  }

  static IllegalArgumentException invalidPayload(JacksonException exception) {
    return GridGrindJsonMessageSupport.invalidPayload(exception);
  }

  static String message(Throwable throwable) {
    return GridGrindJsonMessageSupport.message(throwable);
  }

  static String mismatchedInputMessage(
      tools.jackson.databind.exc.MismatchedInputException exception) {
    return GridGrindJsonMessageSupport.mismatchedInputMessage(exception);
  }

  static String cleanJacksonMessage(String message) {
    return GridGrindJsonMessageSupport.cleanJacksonMessage(message);
  }

  static java.util.Optional<Integer> jsonLine(TokenStreamLocation location) {
    return GridGrindJsonMessageSupport.jsonLine(location);
  }

  static java.util.Optional<Integer> jsonColumn(TokenStreamLocation location) {
    return GridGrindJsonMessageSupport.jsonColumn(location);
  }

  private static WorkbookPlan readRequestTree(JsonNode requestNode) throws IOException {
    return GridGrindJsonCodecSupport.decodeTree(
        requestNode,
        GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER,
        WorkbookPlan.class,
        GridGrindJsonMessageSupport::invalidRequestPayload);
  }

  private static void writeValue(OutputStream outputStream, Object value) throws IOException {
    Objects.requireNonNull(outputStream, "outputStream must not be null");
    Objects.requireNonNull(value, "value must not be null");
    GridGrindJsonMapperSupport.WIRE_JSON_MAPPER.writeValue(outputStream, value);
  }
}
