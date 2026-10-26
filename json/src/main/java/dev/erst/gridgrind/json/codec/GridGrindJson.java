package dev.erst.gridgrind.json.codec;

import dev.erst.gridgrind.core.catalog.Catalog;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.RequestDoctorReport;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.json.intake.GridGrindRequestDecoder;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Objects;

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
  public static RequestIntake analyzeRequest(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return analyzeRequest(inputStream.readAllBytes());
  }

  /** Analyses one UTF-8 request document into bound fragments and every structural problem. */
  public static RequestIntake analyzeRequest(byte[] bytes) {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindRequestDecoder.analyze(bytes);
  }

  /** Reads one explicit authority document using the supplied root for its relative paths. */
  public static ExecutionGrant readGrant(byte[] bytes, Path root) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    Objects.requireNonNull(root, "root must not be null");
    return Objects.requireNonNull(
        GridGrindJsonMapperSupport.JSON_MAPPER
            .readerFor(ExecutionGrant.class)
            .withAttribute(GrantPathDeserializer.ROOT_ATTRIBUTE, root.toAbsolutePath().normalize())
            .readValue(bytes),
        "Grant document must not be null");
  }

  /** Reads a response from an input stream without closing the caller-owned stream. */
  public static WorkbookResult readWorkbookResult(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readValue(
        inputStream,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        WorkbookResult.class,
        JsonReadFailures::syntax);
  }

  /** Reads a response from a byte array. */
  public static WorkbookResult readWorkbookResult(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.readValue(
        bytes,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        WorkbookResult.class,
        JsonReadFailures::syntax);
  }

  /** Reads a protocol catalog from an input stream without closing the caller-owned stream. */
  public static Catalog readProtocolCatalog(InputStream inputStream) throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readValue(
        inputStream,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        Catalog.class,
        JsonReadFailures::syntax);
  }

  /** Reads a protocol catalog from a byte array. */
  public static Catalog readProtocolCatalog(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.readValue(
        bytes, GridGrindJsonMapperSupport.JSON_MAPPER, Catalog.class, JsonReadFailures::syntax);
  }

  /** Reads a request doctor report from an input stream without closing the caller-owned stream. */
  public static RequestDoctorReport readRequestDoctorReport(InputStream inputStream)
      throws IOException {
    Objects.requireNonNull(inputStream, "inputStream must not be null");
    return GridGrindJsonCodecSupport.readValue(
        inputStream,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        RequestDoctorReport.class,
        JsonReadFailures::syntax);
  }

  /** Reads a request doctor report from a byte array. */
  public static RequestDoctorReport readRequestDoctorReport(byte[] bytes) throws IOException {
    Objects.requireNonNull(bytes, "bytes must not be null");
    return GridGrindJsonCodecSupport.readValue(
        bytes,
        GridGrindJsonMapperSupport.JSON_MAPPER,
        RequestDoctorReport.class,
        JsonReadFailures::syntax);
  }

  /** Returns the maximum accepted JSON request document length in bytes. */
  public static long maxRequestDocumentBytes() {
    return GridGrindJsonMapperSupport.maxRequestDocumentBytes();
  }

  /** Rejects one request payload length that exceeds the documented transport limit. */
  public static void requireSupportedRequestLength(long lengthBytes) {
    GridGrindJsonMapperSupport.requireSupportedRequestLength(lengthBytes);
  }
}
