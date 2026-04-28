package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.executor.GridGrindProblems;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Focused tests for response-file fallback behavior in {@link CliResponseWriter}. */
class CliResponseWriterTest extends GridGrindCliTestSupport {
  private final CliResponseWriter responseWriter = new CliResponseWriter();

  @Test
  void writePayloadFallsBackToStructuredFailureWhenTheResponsePathCannotBeWritten()
      throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-payload-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writePayload(
            responseDirectory, stdout, "{\"status\":\"ok\"}".getBytes(StandardCharsets.UTF_8), 0);

    GridGrindResponse.Failure fallback =
        assertInstanceOf(
            GridGrindResponse.Failure.class, GridGrindJson.readResponse(stdout.toByteArray()));

    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.IO_ERROR, fallback.problem().code());
    assertEquals("WRITE_RESPONSE", fallback.problem().context().stage());
    assertEquals(
        java.util.Optional.of(responseDirectory.toAbsolutePath().toString()),
        writeResponseContext(fallback).responsePath());
  }

  @Test
  void writePayloadPreservesOneTrailingNewlineWhenPayloadAlreadyEndsWithNewline()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writePayload(
            null, stdout, "{\"status\":\"ok\"}\n".getBytes(StandardCharsets.UTF_8), 0);

    assertEquals(0, exitCode);
    assertEquals("{\"status\":\"ok\"}\n", stdout.toString(StandardCharsets.UTF_8));
  }

  @Test
  void writePayloadAddsOneTrailingNewlineForEmptyPayload() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode = responseWriter.writePayload(null, stdout, new byte[0], 0);

    assertEquals(0, exitCode);
    assertEquals("\n", stdout.toString(StandardCharsets.UTF_8));
  }

  @Test
  void writeDoctorReportFallsBackToStdoutWhenTheResponsePathCannotBeWritten() throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-doctor-report-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    RequestDoctorReport.Summary summary = summary();
    RequestWarning warning = new RequestWarning(0, "step-1", "SET_CELL", "warning");

    int exitCode =
        responseWriter.writeDoctorReport(
            responseDirectory, stdout, RequestDoctorReport.warnings(summary, List.of(warning)));

    RequestDoctorReport fallback = GridGrindJson.readRequestDoctorReport(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertFalse(fallback.valid());
    assertEquals(java.util.Optional.of(summary), fallback.summary());
    assertEquals(List.of(warning), fallback.warnings());
    assertEquals(GridGrindProblemCode.IO_ERROR, fallback.problem().orElseThrow().code());
    assertEquals("WRITE_RESPONSE", fallback.problem().orElseThrow().context().stage());
    assertEquals(
        java.util.Optional.of(responseDirectory.toAbsolutePath().toString()),
        writeResponseContext(fallback).responsePath());
    assertEquals(1, fallback.problem().orElseThrow().causes().size());
  }

  @Test
  void writeDoctorReportPreservesTheOriginalProblemAsASupplementalCauseWhenFallbackIsNeeded()
      throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-doctor-problem-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    RequestDoctorReport.Summary summary = summary();
    GridGrindResponse.Problem originalProblem =
        GridGrindProblems.problem(
            GridGrindProblemCode.INVALID_REQUEST,
            "bad request",
            new dev.erst.gridgrind.contract.dto.ProblemContext.ValidateRequest(
                dev.erst.gridgrind.contract.dto.ProblemContext.RequestShape.known("NEW", "NONE")),
            new IOException("bad request"));

    int exitCode =
        responseWriter.writeDoctorReport(
            responseDirectory,
            stdout,
            RequestDoctorReport.invalid(summary, List.of(), originalProblem));

    RequestDoctorReport fallback = GridGrindJson.readRequestDoctorReport(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertFalse(fallback.valid());
    assertEquals(java.util.Optional.of(summary), fallback.summary());
    assertEquals(GridGrindProblemCode.IO_ERROR, fallback.problem().orElseThrow().code());
    assertTrue(
        fallback.problem().orElseThrow().causes().stream()
            .anyMatch(
                cause ->
                    cause.code() == GridGrindProblemCode.INVALID_REQUEST
                        && "VALIDATE_REQUEST".equals(cause.stage())
                        && cause.message().contains("bad request")));
  }

  private static RequestDoctorReport.Summary summary() {
    return new RequestDoctorReport.Summary(
        "NEW", "NONE", "FULL_XSSF", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 1, 1, 0, 0);
  }
}
