package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.io.IOException;
import java.io.InputStream;
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
    return GridGrindJsonCodecSupport.decodeTree(
        readRequestTree(inputStream),
        GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER,
        WorkbookPlan.class,
        GridGrindJsonProblemMessageSupport::invalidRequestPayload);
  }

  /** Reads a request from a byte array. */
  public static WorkbookPlan readRequest(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.decodeTree(
        readRequestTree(bytes),
        GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER,
        WorkbookPlan.class,
        GridGrindJsonProblemMessageSupport::invalidRequestPayload);
  }

  /** Reads a request from one in-memory JSON string without exposing a checked I/O seam. */
  public static WorkbookPlan readRequest(String json) {
    Objects.requireNonNull(json, "json must not be null");
    return GridGrindJsonCodecSupport.decodeTree(
        GridGrindJsonCodecSupport.readTree(
            json,
            GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER,
            GridGrindJsonProblemMessageSupport::invalidRequestPayload),
        GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER,
        WorkbookPlan.class,
        GridGrindJsonProblemMessageSupport::invalidRequestPayload);
  }

  /** Reads one request JSON tree from an input stream without closing the caller-owned stream. */
  public static JsonNode readRequestTree(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readTree(
        inputStream,
        GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER,
        GridGrindJsonProblemMessageSupport::invalidRequestPayload);
  }

  /** Reads one request JSON tree from a byte array. */
  public static JsonNode readRequestTree(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    GridGrindJsonMapperSupport.requireSupportedRequestLength(bytes.length);
    return GridGrindJsonCodecSupport.readTree(
        bytes,
        GridGrindJsonMapperSupport.REQUEST_JSON_MAPPER,
        GridGrindJsonProblemMessageSupport::invalidRequestPayload);
  }

  /** Reads a response from an input stream without closing the caller-owned stream. */
  public static GridGrindResponse readResponse(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readValue(
        inputStream,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        GridGrindResponse.class,
        GridGrindJsonProblemMessageSupport::invalidPayload);
  }

  /** Reads a response from a byte array. */
  public static GridGrindResponse readResponse(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.readValue(
        bytes,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        GridGrindResponse.class,
        GridGrindJsonProblemMessageSupport::invalidPayload);
  }

  /** Reads a protocol catalog from an input stream without closing the caller-owned stream. */
  public static Catalog readProtocolCatalog(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readValue(
        inputStream,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        Catalog.class,
        GridGrindJsonProblemMessageSupport::invalidPayload);
  }

  /** Reads a protocol catalog from a byte array. */
  public static Catalog readProtocolCatalog(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.readValue(
        bytes,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        Catalog.class,
        GridGrindJsonProblemMessageSupport::invalidPayload);
  }

  /** Reads a request doctor report from an input stream without closing the caller-owned stream. */
  public static RequestDoctorReport readRequestDoctorReport(InputStream inputStream)
      throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readValue(
        inputStream,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        RequestDoctorReport.class,
        GridGrindJsonProblemMessageSupport::invalidPayload);
  }

  /** Reads a request doctor report from a byte array. */
  public static RequestDoctorReport readRequestDoctorReport(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.readValue(
        bytes,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        RequestDoctorReport.class,
        GridGrindJsonProblemMessageSupport::invalidPayload);
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
    return GridGrindJsonProblemMessageSupport.invalidPayload(exception);
  }

  static String message(Throwable throwable) {
    return GridGrindJsonProblemMessageSupport.message(throwable);
  }

  static String mismatchedInputMessage(
      tools.jackson.databind.exc.MismatchedInputException exception) {
    return GridGrindJsonValueProblemSupport.mismatchedInputMessage(exception);
  }

  static String cleanJacksonMessage(String message) {
    return GridGrindJsonProblemMessageSupport.cleanJacksonMessage(message);
  }

  static java.util.Optional<Integer> jsonLine(TokenStreamLocation location) {
    return GridGrindJsonPayloadMetadataSupport.jsonLine(location);
  }

  static java.util.Optional<Integer> jsonColumn(TokenStreamLocation location) {
    return GridGrindJsonPayloadMetadataSupport.jsonColumn(location);
  }
}
