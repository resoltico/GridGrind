package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.AccessDeniedException;
import java.nio.file.FileSystemException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused tests for response-file fallback behavior in {@link CliResponseWriter}. */
class CliResponseWriterTest extends GridGrindCliTestSupport {
  private final CliResponseWriter responseWriter = new CliResponseWriter();

  @Test
  void writePayloadFallsBackToCliFailureWhenTheResponsePathCannotBeWritten() throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-payload-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writePayload(
            "print-request-template",
            "request template",
            Optional.of("gridgrind --print-request-template"),
            Optional.of(responseDirectory),
            stdout,
            stderr,
            "{\"status\":\"ok\"}".getBytes(StandardCharsets.UTF_8),
            0);

    CliFailureReport fallback = cliFailure(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertEquals(
        "Could not write response file "
            + responseDirectory.toAbsolutePath()
            + ": Is a directory. Wrote the request template to stdout instead."
            + System.lineSeparator(),
        stderr.toString(StandardCharsets.UTF_8));
    assertEquals(GridGrindProblemCode.IO_ERROR, fallback.code());
    assertEquals("print-request-template", fallback.command());
    assertEquals(Optional.of("--response"), fallback.argument());
    assertTrue(
        fallback
            .message()
            .startsWith("Could not write response file " + responseDirectory.toAbsolutePath()));
    assertEquals(
        Optional.of(
            "Provide one writable file path after --response, or remove --response to print the request template on stdout."),
        fallback.resolution());
  }

  @Test
  void writePayloadPreservesOneTrailingNewlineWhenPayloadAlreadyEndsWithNewline()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writePayload(
            "help",
            "help text",
            Optional.of("gridgrind --help"),
            Optional.empty(),
            stdout,
            "{\"status\":\"ok\"}\n".getBytes(StandardCharsets.UTF_8),
            0);

    assertEquals(0, exitCode);
    assertEquals("{\"status\":\"ok\"}\n", stdout.toString(StandardCharsets.UTF_8));
  }

  @Test
  void writePayloadAddsOneTrailingNewlineForEmptyPayload() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writePayload(
            "help",
            "help text",
            Optional.of("gridgrind --help"),
            Optional.empty(),
            stdout,
            new byte[0],
            0);

    assertEquals(0, exitCode);
    assertEquals("\n", stdout.toString(StandardCharsets.UTF_8));
  }

  @Test
  void writePayloadRoutesNonSuccessPayloadsToStdoutWhenNoResponsePathIsConfigured()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writePayload(
            "print-task-keyword-match",
            "task keyword match report",
            Optional.of("gridgrind --print-task-keyword-match --query \"monthly sales dashboard\""),
            Optional.empty(),
            stdout,
            stderr,
            "{\"status\":\"error\"}".getBytes(StandardCharsets.UTF_8),
            2);

    assertEquals(2, exitCode);
    assertEquals("{\"status\":\"error\"}\n", stdout.toString(StandardCharsets.UTF_8));
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
  }

  @Test
  void writeCliFailureReportWithoutResponsePathWritesTheCompactFailureToStdout()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writeCliFailureReport(
            Optional.empty(),
            stdout,
            stderr,
            new dev.erst.gridgrind.cli.discovery.CliFailureReport(
                GridGrindProtocolVersion.current(),
                2,
                "parse-arguments",
                GridGrindProblemCode.INVALID_ARGUMENTS,
                "bad flag",
                dev.erst.gridgrind.cli.discovery.CliFailureLocation.unavailable(),
                Optional.of("--flag"),
                List.of("gridgrind --help"),
                Optional.of("Use one primary command.")));

    CliFailureReport failure = cliFailureOnStdout(stdout, stderr);
    assertEquals(2, exitCode);
    assertEquals("bad flag", failure.message());
  }

  @Test
  void writeCliFailureReportFallsBackToTheCliFailureReportWhenTheResponsePathCannotBeWritten()
      throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-cli-failure-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writeCliFailureReport(
            Optional.of(responseDirectory),
            stdout,
            stderr,
            new dev.erst.gridgrind.cli.discovery.CliFailureReport(
                GridGrindProtocolVersion.current(),
                2,
                "parse-arguments",
                GridGrindProblemCode.INVALID_ARGUMENTS,
                "bad flag",
                dev.erst.gridgrind.cli.discovery.CliFailureLocation.unavailable(),
                Optional.of("--flag"),
                List.of("gridgrind --help"),
                Optional.of("Use one primary command.")));

    CliFailureReport fallback = cliFailure(stdout.toByteArray());

    assertEquals(2, exitCode);
    assertEquals(
        "Could not write response file "
            + responseDirectory.toAbsolutePath()
            + ": Is a directory. Wrote the CLI failure report to stdout instead."
            + System.lineSeparator(),
        stderr.toString(StandardCharsets.UTF_8));
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, fallback.code());
    assertEquals("parse-arguments", fallback.command());
    assertEquals(Optional.of("--flag"), fallback.argument());
  }

  @Test
  void writeRequestFailureReportEmitsRequestFailurePointerOnStderr() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-request-failure-report-", ".json");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writeRequestFailureReport(
            Optional.of(responsePath),
            stdout,
            stderr,
            new dev.erst.gridgrind.cli.discovery.CliFailureReport(
                GridGrindProtocolVersion.current(),
                1,
                "execute",
                GridGrindProblemCode.INVALID_REQUEST_SHAPE,
                "missing required field",
                dev.erst.gridgrind.cli.discovery.CliFailureLocation.unavailable(),
                Optional.of("--request"),
                List.of("gridgrind --help-protocol"),
                Optional.empty()));

    assertEquals(1, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertTrue(
        stderr
            .toString(StandardCharsets.UTF_8)
            .contains(
                "GridGrind wrote the request failure report to " + responsePath.toAbsolutePath()));
    assertTrue(stderr.toString(StandardCharsets.UTF_8).contains("[INVALID_REQUEST_SHAPE:"));
    CliFailureReport failure = cliFailure(Files.readAllBytes(responsePath));
    assertEquals(GridGrindProblemCode.INVALID_REQUEST_SHAPE, failure.code());
  }

  @Test
  void writeDoctorReportFallsBackToStdoutWhenTheResponsePathCannotBeWritten() throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-doctor-report-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    RequestDoctorReport.Summary summary = summary();
    RequestWarning warning = new RequestWarning(0, "step-1", "SET_CELL", "warning");

    int exitCode =
        responseWriter.writeDoctorReport(
            Optional.of(responseDirectory),
            stdout,
            stderr,
            RequestDoctorReport.warnings(summary, List.of(warning)));

    RequestDoctorReport fallback = GridGrindJson.readRequestDoctorReport(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertEquals(
        "Could not write response file "
            + responseDirectory.toAbsolutePath()
            + ": Is a directory. Wrote the doctor report to stdout instead."
            + System.lineSeparator(),
        stderr.toString(StandardCharsets.UTF_8));
    assertFalse(fallback.valid());
    assertEquals(java.util.Optional.of(summary), fallback.summary());
    assertEquals(List.of(warning), fallback.warnings());
    assertEquals(GridGrindProblemCode.IO_ERROR, fallback.primaryProblem().orElseThrow().code());
    assertEquals("WRITE_RESPONSE", fallback.primaryProblem().orElseThrow().context().stage());
    assertTrue(
        fallback
            .primaryProblem()
            .orElseThrow()
            .message()
            .startsWith("Could not write response file " + responseDirectory.toAbsolutePath()),
        "doctor fallback must explain that response writing failed");
    assertEquals(
        java.util.Optional.of(responseDirectory.toAbsolutePath().toString()),
        writeResponseContext(fallback).responsePath());
    assertEquals(1, fallback.primaryProblem().orElseThrow().causes().size());
  }

  @Test
  void writeToResponseFileEmitsStderrPointerForNonSuccessResponses() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-failure-response-", ".json");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    GridGrindResponse.Failure failure =
        GridGrindResponses.failure(
            dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
            GridGrindProblems.problem(
                GridGrindProblemCode.INVALID_REQUEST,
                "bad request",
                new dev.erst.gridgrind.contract.dto.ProblemContext.ValidateRequest(
                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape
                        .known("NEW", "NONE")),
                new IllegalArgumentException("bad request")));

    int exitCode = responseWriter.write(Optional.of(responsePath), stdout, stderr, failure);

    assertEquals(1, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertTrue(
        stderr
            .toString(StandardCharsets.UTF_8)
            .contains("GridGrind wrote the response to " + responsePath.toAbsolutePath()));
    assertTrue(stderr.toString(StandardCharsets.UTF_8).contains("[INVALID_REQUEST: bad request]"));
    assertInstanceOf(
        GridGrindResponse.Failure.class,
        GridGrindJson.readResponse(Files.readAllBytes(responsePath)));
  }

  @Test
  void writeWithExplicitLogicalExitCodeDelegatesToTheSharedResponsePathFlow() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-explicit-exit-response-", ".json");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.write(
            Optional.of(responsePath),
            stdout,
            GridGrindResponses.success(
                java.util.List.of(), java.util.List.of(), java.util.List.of()),
            2);

    assertEquals(2, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertInstanceOf(
        GridGrindResponse.Success.class,
        GridGrindJson.readResponse(Files.readAllBytes(responsePath)));
  }

  @Test
  void writeWithExplicitLogicalExitCodeEmitsCompactPointerForSuccessPayloads() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-explicit-stderr-response-", ".json");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.write(
            Optional.of(responsePath),
            stdout,
            stderr,
            GridGrindResponses.success(
                java.util.List.of(), java.util.List.of(), java.util.List.of()),
            2);

    assertEquals(2, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertEquals(
        "GridGrind wrote the response to "
            + responsePath.toAbsolutePath()
            + "; inspect that file for failure."
            + System.lineSeparator(),
        stderr.toString(StandardCharsets.UTF_8));
  }

  @Test
  void formattedProblemSummaryOmitsBlankTextAndFormatsVisibleText() {
    assertEquals("", CliResponseWriter.formattedProblemSummary(Optional.empty()));
    assertEquals("", CliResponseWriter.formattedProblemSummary(Optional.of(" ")));
    assertEquals(
        " [INVALID_REQUEST: bad request]",
        CliResponseWriter.formattedProblemSummary(Optional.of("INVALID_REQUEST: bad request")));
  }

  @Test
  void writeWithoutExplicitStderrDelegatesToTheSharedResponsePathFlow() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-default-stderr-response-", ".json");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.write(
            Optional.of(responsePath),
            stdout,
            GridGrindResponses.success(
                java.util.List.of(), java.util.List.of(), java.util.List.of()));

    assertEquals(0, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertInstanceOf(
        GridGrindResponse.Success.class,
        GridGrindJson.readResponse(Files.readAllBytes(responsePath)));
  }

  @Test
  void writeDoctorReportToResponseFileEmitsStderrPointerForInvalidReports() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-invalid-doctor-report-", ".json");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    RequestDoctorReport report =
        RequestDoctorReport.invalid(
            summary(),
            List.of(),
            GridGrindProblems.problem(
                GridGrindProblemCode.INVALID_REQUEST,
                "bad request",
                new dev.erst.gridgrind.contract.dto.ProblemContext.ValidateRequest(
                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape
                        .known("NEW", "NONE")),
                new IllegalArgumentException("bad request")));

    int exitCode =
        responseWriter.writeDoctorReport(Optional.of(responsePath), stdout, stderr, report);

    assertEquals(1, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertTrue(
        stderr
            .toString(StandardCharsets.UTF_8)
            .contains("GridGrind wrote the doctor report to " + responsePath.toAbsolutePath()));
    assertTrue(stderr.toString(StandardCharsets.UTF_8).contains("[INVALID_REQUEST: bad request]"));
    assertFalse(GridGrindJson.readRequestDoctorReport(Files.readAllBytes(responsePath)).valid());
  }

  @Test
  void writeDoctorReportPreservesTheOriginalProblemAsASupplementalCauseWhenFallbackIsNeeded()
      throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-doctor-problem-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    RequestDoctorReport.Summary summary = summary();
    GridGrindProblemDetail.Problem originalProblem =
        GridGrindProblems.problem(
            GridGrindProblemCode.INVALID_REQUEST,
            "bad request",
            new dev.erst.gridgrind.contract.dto.ProblemContext.ValidateRequest(
                dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape.known(
                    "NEW", "NONE")),
            new IOException("bad request"));

    int exitCode =
        responseWriter.writeDoctorReport(
            Optional.of(responseDirectory),
            stdout,
            RequestDoctorReport.invalid(summary, List.of(), originalProblem));

    RequestDoctorReport fallback = GridGrindJson.readRequestDoctorReport(stdout.toByteArray());

    assertEquals(1, exitCode);
    assertFalse(fallback.valid());
    assertEquals(java.util.Optional.of(summary), fallback.summary());
    assertEquals(GridGrindProblemCode.IO_ERROR, fallback.primaryProblem().orElseThrow().code());
    assertTrue(
        fallback.primaryProblem().orElseThrow().causes().stream()
            .anyMatch(
                cause ->
                    cause.code() == GridGrindProblemCode.INVALID_REQUEST
                        && "VALIDATE_REQUEST".equals(cause.stage())
                        && cause.message().contains("bad request")));
  }

  @Test
  void writeResponseProblemFormatsAccessDeniedFailuresWithPermissionMessage() {
    Path responsePath = Path.of("/tmp/response.json");

    GridGrindProblemDetail.Problem problem =
        CliResponseWriter.writeResponseProblem(
            new AccessDeniedException(responsePath.toString()), responsePath);

    assertEquals(
        "Could not write response file /tmp/response.json: permission denied", problem.message());
    assertEquals(problem.message(), problem.causes().getFirst().message());
  }

  @Test
  void writeResponseProblemFormatsFileSystemReasonWhenAvailable() {
    Path responsePath = Path.of("/tmp/response.json");

    GridGrindProblemDetail.Problem problem =
        CliResponseWriter.writeResponseProblem(
            new FileSystemException(responsePath.toString(), null, "Is a directory"), responsePath);

    assertEquals(
        "Could not write response file /tmp/response.json: Is a directory", problem.message());
    assertEquals(problem.message(), problem.causes().getFirst().message());
  }

  @Test
  void writeResponseProblemFormatsOtherFileConflictsWhenReasonIsMissing() {
    Path responsePath = Path.of("/tmp/response.json");

    GridGrindProblemDetail.Problem problem =
        CliResponseWriter.writeResponseProblem(
            new FileSystemException(responsePath.toString(), "/tmp/other.json", null),
            responsePath);

    assertEquals(
        "Could not write response file /tmp/response.json: conflict with /tmp/other.json",
        problem.message());
    assertEquals(problem.message(), problem.causes().getFirst().message());
  }

  @Test
  void writeResponseProblemFormatsGenericIoFailuresWithoutRawPathCollapse() {
    Path responsePath = Path.of("/tmp/response.json");

    GridGrindProblemDetail.Problem blankMessageProblem =
        CliResponseWriter.writeResponseProblem(new IOException(), responsePath);
    GridGrindProblemDetail.Problem blankStringProblem =
        CliResponseWriter.writeResponseProblem(new IOException("   "), responsePath);
    GridGrindProblemDetail.Problem explicitMessageProblem =
        CliResponseWriter.writeResponseProblem(new IOException("disk full"), responsePath);

    assertEquals("Could not write response file /tmp/response.json", blankMessageProblem.message());
    assertEquals("Could not write response file /tmp/response.json", blankStringProblem.message());
    assertEquals(
        "Could not write response file /tmp/response.json: disk full",
        explicitMessageProblem.message());
  }

  @Test
  void writeResponseProblemFallsBackWhenFileSystemReasonAndOtherFileAreBlank() {
    Path responsePath = Path.of("/tmp/response.json");

    GridGrindProblemDetail.Problem problem =
        CliResponseWriter.writeResponseProblem(
            new FileSystemException(responsePath.toString(), "   ", "   "), responsePath);

    assertEquals("Could not write response file /tmp/response.json", problem.message());
    assertEquals(problem.message(), problem.causes().getFirst().message());
  }

  @Test
  void writeResponseProblemFallsBackWhenFileSystemReasonAndOtherFileAreMissing() {
    Path responsePath = Path.of("/tmp/response.json");

    GridGrindProblemDetail.Problem problem =
        CliResponseWriter.writeResponseProblem(
            new FileSystemException(responsePath.toString(), null, null), responsePath);

    assertEquals("Could not write response file /tmp/response.json", problem.message());
    assertEquals(problem.message(), problem.causes().getFirst().message());
  }

  private static RequestDoctorReport.Summary summary() {
    return new RequestDoctorReport.Summary(
        "NEW", "NONE", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 1, 1, 0, 0);
  }
}
