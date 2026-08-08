package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.cli.discovery.CommandError;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResults;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.AccessDeniedException;
import java.nio.file.FileAlreadyExistsException;
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
  void requestDiagnosticOutputBoundaryRedactsOnlyTheProblemAtItsSecretOwnerPath()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    String rawRequest =
        """
        {
          "protocolVersion":"V2",
          "source":{"type":"EXISTING","path":"source.xlsx","security":{"password":"source-secret"}},
          "persistence":{"type":"NONE"},
          "steps":[]
        }
        """;
    CommandError diagnostic =
        new CommandError(
            GridGrindProtocolVersion.current(),
            1,
            "execute",
            List.of(),
            List.of(
                GridGrindProblemDetail.Problem.of(
                    GridGrindProblemCode.INVALID_REQUEST,
                    "source-secret",
                    new ProblemContext.ReadRequest(
                        dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput
                            .standardInput(),
                        dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation
                            .pathOnly("source.security.password")))),
            Optional.empty());

    int exitCode =
        responseWriter.writeCommandError(
            Optional.empty(),
            stdout,
            stderr,
            diagnostic,
            GridGrindJson.analyzeRequest(rawRequest.getBytes(StandardCharsets.UTF_8))
                .diagnosticRedactor(),
            false);

    String rendered = stderr.toString(StandardCharsets.UTF_8);
    assertEquals(1, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertFalse(rendered.contains("source-secret"));
    assertEquals("[REDACTED]", commandErrorOnStdout(stderr).primaryProblem().message());
  }

  @Test
  void writePayloadFallsBackToCommandErrorWhenTheResponsePathCannotBeWritten() throws IOException {
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
            0,
            false);

    CommandError fallback = commandError(stdout.toByteArray());
    CommandError stderrDiagnostic = commandErrorOnStdout(stderr);

    assertEquals(1, exitCode);
    assertEquals(fallback, stderrDiagnostic);
    assertEquals(GridGrindProblemCode.IO_ERROR, fallback.primaryProblem().code());
    assertEquals("print-request-template", fallback.command());
    assertEquals(
        Optional.of(responseDirectory.toAbsolutePath().toString()),
        writeResponseContext(fallback).responsePath());
    assertEquals(
        Optional.of("STDOUT"),
        fallback
            .transport()
            .map(
                transport ->
                    switch (transport) {
                      case dev.erst.gridgrind.cli.discovery.CliTransport.StandardOutput _ ->
                          "STDOUT";
                      case dev.erst.gridgrind.cli.discovery.CliTransport.ResponseFile _ -> "FILE";
                    }));
    assertTrue(
        fallback
            .primaryProblem()
            .message()
            .startsWith("Could not write response file " + responseDirectory.toAbsolutePath()));
    assertEquals(
        "Check the --response destination path, parent directory permissions, free disk space, and file locks before retrying.",
        fallback.primaryProblem().resolution());
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
            OutputStream.nullOutputStream(),
            "{\"status\":\"ok\"}\n".getBytes(StandardCharsets.UTF_8),
            0,
            false);

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
            OutputStream.nullOutputStream(),
            new byte[0],
            0,
            false);

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
            "print-recipe-keyword-match",
            "task keyword match report",
            Optional.of(
                "gridgrind --print-recipe-keyword-match --query \"monthly sales dashboard\""),
            Optional.empty(),
            stdout,
            stderr,
            "{\"status\":\"error\"}".getBytes(StandardCharsets.UTF_8),
            2,
            false);

    assertEquals(2, exitCode);
    assertEquals("{\"status\":\"error\"}\n", stdout.toString(StandardCharsets.UTF_8));
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
  }

  @Test
  void writeCommandErrorWithoutResponsePathWritesTheCompactDiagnosticToStderr()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writeCommandError(
            Optional.empty(),
            stdout,
            stderr,
            new dev.erst.gridgrind.cli.discovery.CommandError(
                GridGrindProtocolVersion.current(),
                2,
                "cli",
                List.of("gridgrind --help"),
                List.of(
                    GridGrindProblemDetail.Problem.of(
                        GridGrindProblemCode.INVALID_ARGUMENTS,
                        "bad flag",
                        new ProblemContext.ParseArguments(CliArgument.named("--flag")))),
                Optional.empty()),
            false);

    CommandError failure = commandErrorOnStdout(stdout, stderr);
    assertEquals(2, exitCode);
    assertEquals("bad flag", failure.primaryProblem().message());
  }

  @Test
  void writeCommandErrorFallsBackToTheCommandErrorWhenTheResponsePathCannotBeWritten()
      throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-cli-failure-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writeCommandError(
            Optional.of(responseDirectory),
            stdout,
            stderr,
            new dev.erst.gridgrind.cli.discovery.CommandError(
                GridGrindProtocolVersion.current(),
                2,
                "cli",
                List.of("gridgrind --help"),
                List.of(
                    GridGrindProblemDetail.Problem.of(
                        GridGrindProblemCode.INVALID_ARGUMENTS,
                        "bad flag",
                        new ProblemContext.ParseArguments(CliArgument.named("--flag")))),
                Optional.empty()),
            false);

    CommandError fallback = commandError(stdout.toByteArray());
    CommandError stderrDiagnostic = commandErrorOnStdout(stderr);

    assertEquals(2, exitCode);
    assertEquals(fallback, stderrDiagnostic);
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, fallback.primaryProblem().code());
    assertEquals("cli", fallback.command());
    assertEquals(Optional.of("--flag"), parseArgumentsContext(fallback).argumentName());
    assertEquals(
        Optional.of("STDOUT"),
        fallback
            .transport()
            .map(
                transport ->
                    switch (transport) {
                      case dev.erst.gridgrind.cli.discovery.CliTransport.StandardOutput _ ->
                          "STDOUT";
                      case dev.erst.gridgrind.cli.discovery.CliTransport.ResponseFile _ -> "FILE";
                    }));
  }

  @Test
  void writeRequestFailureReportMirrorsThePersistedDiagnosticOnStderr() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-request-failure-report-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writeCommandError(
            Optional.of(responsePath),
            stdout,
            stderr,
            new dev.erst.gridgrind.cli.discovery.CommandError(
                GridGrindProtocolVersion.current(),
                1,
                "execute",
                List.of("gridgrind --help-protocol"),
                List.of(
                    GridGrindProblemDetail.Problem.of(
                        GridGrindProblemCode.INVALID_REQUEST_SHAPE,
                        "missing required field",
                        new ProblemContext.ReadRequest(
                            dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces
                                .RequestInput.standardInput(),
                            dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces
                                .JsonLocation.pathOnly("steps[0].type")))),
                Optional.empty()),
            false);

    assertEquals(1, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    CommandError failure = commandError(Files.readAllBytes(responsePath));
    CommandError stderrDiagnostic = commandErrorOnStdout(stderr);
    assertEquals(failure, stderrDiagnostic);
    assertEquals(GridGrindProblemCode.INVALID_REQUEST_SHAPE, failure.primaryProblem().code());
    assertEquals(
        Optional.of(responsePath.toAbsolutePath().toString()),
        failure.transport().flatMap(transport -> transport.responsePathValue()));
  }

  @Test
  void writeDoctorReportFallsBackToStdoutWhenTheResponsePathCannotBeWritten() throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-doctor-report-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    RequestDoctorReport.Summary summary = summary();
    RequestWarning warning = new RequestWarning(dev.erst.gridgrind.contract.dto.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA, 0, "step-1", "SET_CELL", "warning");

    int exitCode =
        responseWriter.writeDoctorReport(
            Optional.of(responseDirectory),
            stdout,
            stderr,
            RequestDoctorReport.warnings(summary, List.of(warning)),
            false);

    RequestDoctorReport fallback = GridGrindJson.readRequestDoctorReport(stdout.toByteArray());
    CommandError stderrDiagnostic = commandErrorOnStdout(stderr);

    assertEquals(1, exitCode);
    assertFalse(fallback.valid());
    assertEquals(java.util.Optional.of(summary), fallback.summary());
    assertEquals(List.of(warning), fallback.warnings());
    assertEquals(GridGrindProblemCode.IO_ERROR, fallback.primaryProblem().orElseThrow().code());
    assertEquals(fallback.primaryProblem().orElseThrow(), stderrDiagnostic.primaryProblem());
    assertEquals(
        Optional.of("STDOUT"),
        stderrDiagnostic
            .transport()
            .map(
                transport ->
                    switch (transport) {
                      case dev.erst.gridgrind.cli.discovery.CliTransport.StandardOutput _ ->
                          "STDOUT";
                      case dev.erst.gridgrind.cli.discovery.CliTransport.ResponseFile _ -> "FILE";
                    }));
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
  void writeDoctorReportDoesNotEmitStderrWhenOneValidReportWasPersisted() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-valid-doctor-report-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    RequestDoctorReport report =
        RequestDoctorReport.warnings(
            summary(), List.of(new RequestWarning(dev.erst.gridgrind.contract.dto.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA, 0, "step-1", "SET_CELL", "warning")));

    int exitCode =
        responseWriter.writeDoctorReport(Optional.of(responsePath), stdout, stderr, report, false);

    assertEquals(0, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
    assertEquals(report, GridGrindJson.readRequestDoctorReport(Files.readAllBytes(responsePath)));
  }

  @Test
  void writeToResponseFileEmitsCommandErrorOnStderrForNonSuccessResponses() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-failure-response-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    WorkbookResult.Failure failure =
        WorkbookResults.failure(
            dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
            GridGrindProblems.problem(
                GridGrindProblemCode.INVALID_REQUEST,
                "bad request",
                new dev.erst.gridgrind.contract.dto.ProblemContext.ValidateRequest(
                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape
                        .known("NEW", "NONE")),
                new IllegalArgumentException("bad request")));

    int exitCode =
        responseWriter.write(
            Optional.of(responsePath),
            stdout,
            stderr,
            failure,
            CliResponseTransportSupport.exitCodeFor(failure),
            false);

    CommandError stderrDiagnostic = commandErrorOnStdout(stderr);

    assertEquals(1, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    WorkbookResult.Failure persistedFailure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            GridGrindJson.readWorkbookResult(Files.readAllBytes(responsePath)));
    assertEquals(persistedFailure.primaryProblem(), stderrDiagnostic.primaryProblem());
    assertEquals(
        Optional.of(responsePath.toAbsolutePath().toString()),
        stderrDiagnostic.transport().flatMap(transport -> transport.responsePathValue()));
  }

  @Test
  void writeWithExplicitLogicalExitCodeDelegatesToTheSharedResponsePathFlow() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-explicit-exit-response-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.write(
            Optional.of(responsePath),
            stdout,
            OutputStream.nullOutputStream(),
            WorkbookResults.success(
                java.util.List.of(), java.util.List.of(), java.util.List.of()),
            2,
            false);

    assertEquals(2, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertInstanceOf(
        WorkbookResult.Success.class,
        GridGrindJson.readWorkbookResult(Files.readAllBytes(responsePath)));
  }

  @Test
  void writeWithExplicitLogicalExitCodeDoesNotInventDiagnosticsForSuccessPayloads()
      throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-explicit-stderr-response-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.write(
            Optional.of(responsePath),
            stdout,
            stderr,
            WorkbookResults.success(
                java.util.List.of(), java.util.List.of(), java.util.List.of()),
            2,
            false);

    assertEquals(2, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
  }

  @Test
  void writeWithoutExplicitStderrDelegatesToTheSharedResponsePathFlow() throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-default-stderr-response-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.write(
            Optional.of(responsePath),
            stdout,
            OutputStream.nullOutputStream(),
            WorkbookResults.success(
                java.util.List.of(), java.util.List.of(), java.util.List.of()),
            0,
            false);

    assertEquals(0, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertInstanceOf(
        WorkbookResult.Success.class,
        GridGrindJson.readWorkbookResult(Files.readAllBytes(responsePath)));
  }

  @Test
  void writeDoctorReportToResponseFileEmitsCommandErrorOnStderrForInvalidReports()
      throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-invalid-doctor-report-", ".json");
    Files.deleteIfExists(responsePath);
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
        responseWriter.writeDoctorReport(Optional.of(responsePath), stdout, stderr, report, false);

    CommandError stderrDiagnostic = commandErrorOnStdout(stderr);

    assertEquals(1, exitCode);
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    RequestDoctorReport persistedReport =
        GridGrindJson.readRequestDoctorReport(Files.readAllBytes(responsePath));
    assertFalse(persistedReport.valid());
    assertEquals(persistedReport.primaryProblem().orElseThrow(), stderrDiagnostic.primaryProblem());
    assertEquals(
        Optional.of(responsePath.toAbsolutePath().toString()),
        stderrDiagnostic.transport().flatMap(transport -> transport.responsePathValue()));
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
            OutputStream.nullOutputStream(),
            RequestDoctorReport.invalid(summary, List.of(), originalProblem),
            false);

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
  void writeResponseFallbackMirrorsTheStdoutProblemOnStderrAndPreservesCauseChain()
      throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-response-problem-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    WorkbookResult.Failure originalFailure =
        WorkbookResults.failure(
            dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
            GridGrindProblems.problem(
                GridGrindProblemCode.IO_ERROR,
                "save failed",
                new dev.erst.gridgrind.contract.dto.ProblemContext.PersistWorkbook(
                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape
                        .known("EXISTING", "SAVE_AS"),
                    Optional.empty(),
                    Optional.of("/tmp/output.xlsx")),
                new IOException("save failed")));

    int exitCode =
        responseWriter.write(
            Optional.of(responseDirectory),
            stdout,
            stderr,
            originalFailure,
            CliResponseTransportSupport.exitCodeFor(originalFailure),
            false);

    WorkbookResult.Failure fallbackResponse =
        assertInstanceOf(
            WorkbookResult.Failure.class, GridGrindJson.readWorkbookResult(stdout.toByteArray()));
    CommandError stderrDiagnostic = commandErrorOnStdout(stderr);

    assertEquals(1, exitCode);
    assertEquals(fallbackResponse.primaryProblem(), stderrDiagnostic.primaryProblem());
    assertEquals(GridGrindProblemCode.IO_ERROR, stderrDiagnostic.primaryProblem().code());
    assertEquals(
        List.of("WRITE_RESPONSE", "PERSIST_WORKBOOK"),
        stderrDiagnostic.primaryProblem().causes().stream()
            .map(GridGrindProblemDetail.ProblemCause::stage)
            .toList());
    assertEquals(
        Optional.of("STDOUT"),
        stderrDiagnostic
            .transport()
            .map(
                transport ->
                    switch (transport) {
                      case dev.erst.gridgrind.cli.discovery.CliTransport.StandardOutput _ ->
                          "STDOUT";
                      case dev.erst.gridgrind.cli.discovery.CliTransport.ResponseFile _ -> "FILE";
                    }));
  }

  @Test
  void writeResponseProblemFormatsAccessDeniedFailuresWithPermissionMessage() {
    Path responsePath = Path.of("/tmp/response.json");

    GridGrindProblemDetail.Problem problem =
        CliResponseTransportSupport.writeResponseProblem(
            new AccessDeniedException(responsePath.toString()), responsePath);

    assertEquals(
        "Could not write response file /tmp/response.json: permission denied", problem.message());
    assertEquals(problem.message(), problem.causes().getFirst().message());
  }

  @Test
  void writeResponseProblemFormatsFileSystemReasonWhenAvailable() {
    Path responsePath = Path.of("/tmp/response.json");

    GridGrindProblemDetail.Problem problem =
        CliResponseTransportSupport.writeResponseProblem(
            new FileSystemException(responsePath.toString(), null, "Is a directory"), responsePath);

    assertEquals(
        "Could not write response file /tmp/response.json: Is a directory", problem.message());
    assertEquals(problem.message(), problem.causes().getFirst().message());
  }

  @Test
  void writeResponseProblemFormatsOtherFileConflictsWhenReasonIsMissing() {
    Path responsePath = Path.of("/tmp/response.json");

    GridGrindProblemDetail.Problem problem =
        CliResponseTransportSupport.writeResponseProblem(
            new FileSystemException(responsePath.toString(), "/tmp/other.json", null),
            responsePath);

    assertEquals(
        "Could not write response file /tmp/response.json: conflict with /tmp/other.json",
        problem.message());
    assertEquals(problem.message(), problem.causes().getFirst().message());
  }

  @Test
  void responseWriteMessageFormatsExistingFileAndDirectoryConflicts() throws IOException {
    Path existingDirectory = Files.createTempDirectory("gridgrind-response-writer-existing-dir-");
    Path existingFile = Files.createTempFile("gridgrind-response-writer-existing-file-", ".json");

    assertEquals(
        "Could not write response file " + existingDirectory + ": Is a directory",
        CliResponseTransportSupport.responseWriteMessage(
            new FileAlreadyExistsException(existingDirectory.toString()), existingDirectory));
    assertEquals(
        "Could not write response file "
            + existingFile
            + ": already exists; GridGrind never replaces an existing response file implicitly",
        CliResponseTransportSupport.responseWriteMessage(
            new FileAlreadyExistsException(existingFile.toString()), existingFile));
  }

  @Test
  void writeResponseProblemFormatsGenericIoFailuresWithoutRawPathCollapse() {
    Path responsePath = Path.of("/tmp/response.json");

    GridGrindProblemDetail.Problem blankMessageProblem =
        CliResponseTransportSupport.writeResponseProblem(new IOException(), responsePath);
    GridGrindProblemDetail.Problem blankStringProblem =
        CliResponseTransportSupport.writeResponseProblem(new IOException("   "), responsePath);
    GridGrindProblemDetail.Problem explicitMessageProblem =
        CliResponseTransportSupport.writeResponseProblem(
            new IOException("disk full"), responsePath);

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
        CliResponseTransportSupport.writeResponseProblem(
            new FileSystemException(responsePath.toString(), "   ", "   "), responsePath);

    assertEquals("Could not write response file /tmp/response.json", problem.message());
    assertEquals(problem.message(), problem.causes().getFirst().message());
  }

  @Test
  void writeResponseProblemFallsBackWhenFileSystemReasonAndOtherFileAreMissing() {
    Path responsePath = Path.of("/tmp/response.json");

    GridGrindProblemDetail.Problem problem =
        CliResponseTransportSupport.writeResponseProblem(
            new FileSystemException(responsePath.toString(), null, null), responsePath);

    assertEquals("Could not write response file /tmp/response.json", problem.message());
    assertEquals(problem.message(), problem.causes().getFirst().message());
  }

  private static RequestDoctorReport.Summary summary() {
    return new RequestDoctorReport.Summary(
        "NEW", "NONE", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 1, 1, 0, 0);
  }
}
