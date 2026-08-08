package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import java.io.IOException;
import java.io.InputStream;
import java.util.Objects;
import tools.jackson.core.JacksonException;
import tools.jackson.core.TokenStreamLocation;

/** Shared JSON codec for the GridGrind protocol. */
public final class GridGrindJson {
  private GridGrindJson() {}

  /** Reads a request from an input stream without closing the caller-owned stream. */
  public static WorkbookPlan readRequest(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    byte[] bytes = inputStream.readAllBytes();
    return readRequest(bytes);
  }

  /** Reads a request from a byte array. */
  public static WorkbookPlan readRequest(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindRequestDecoder.read(bytes);
  }

  /** Reads a request from one in-memory JSON string without exposing a checked I/O seam. */
  public static WorkbookPlan readRequest(String json) {
    Objects.requireNonNull(json, "json must not be null");
    return GridGrindRequestDecoder.read(json);
  }

  /** Analyses one request byte stream without discarding valid sibling fragments after a defect. */
  public static RequestAnalysis analyzeRequest(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return analyzeRequest(inputStream.readAllBytes());
  }

  /** Analyses one UTF-8 request document into bound fragments and every structural problem. */
  public static RequestAnalysis analyzeRequest(byte[] bytes) {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindRequestDecoder.analyze(bytes);
  }

  /** Reads a response from an input stream without closing the caller-owned stream. */
  public static WorkbookResult readWorkbookResult(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readValue(
        inputStream,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        WorkbookResult.class,
        GridGrindJsonProblemMessageSupport::invalidPayload);
  }

  /** Reads a response from a byte array. */
  public static WorkbookResult readWorkbookResult(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.readValue(
        bytes,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        WorkbookResult.class,
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

  /** Converts one tolerant-parser finding into a canonical classified exception. */
  public static IllegalArgumentException structuralException(RequestStructuralProblem problem) {
    Objects.requireNonNull(problem, "problem must not be null");
    return GridGrindRequestDecoder.structuralException(problem);
  }
}
