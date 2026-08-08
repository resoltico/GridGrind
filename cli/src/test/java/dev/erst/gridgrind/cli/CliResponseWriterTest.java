package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.cli.discovery.CliTransportNotice;
import dev.erst.gridgrind.cli.discovery.CommandError;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResults;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Contract tests for primary-payload routing and its one stderr transport notice. */
class CliResponseWriterTest extends GridGrindCliTestSupport {
  private final CliResponseWriter responseWriter = new CliResponseWriter();

  @Test
  void rejectedCommandIsTheSoleStdoutPayloadWithoutAResponseFile() throws Exception {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writeCommandError(Optional.empty(), stdout, stderr, commandError(), false);

    CommandError written = commandErrorOnStdout(stdout, stderr);
    assertEquals(2, exitCode);
    assertEquals("REJECTED", written.status());
    assertEquals("execute", written.command());
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, written.primaryProblem().code());
  }

  @Test
  void rejectedCommandResponseFileFailureRetainsProblemsOnStdoutAndAddsOnlyTransportMetadata()
      throws Exception {
    Path responseDirectory = Files.createTempDirectory("gridgrind-command-error-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writeCommandError(
            Optional.of(responseDirectory), stdout, stderr, commandError(), false);

    CommandError fallback = commandError(stdout.toByteArray());
    CliTransportNotice notice =
        GridGrindCliJson.readBytes(stderr.toByteArray(), CliTransportNotice.class);
    assertEquals(1, exitCode);
    assertEquals(2, fallback.problems().size());
    assertEquals(GridGrindProblemCode.INVALID_ARGUMENTS, fallback.problems().getFirst().code());
    assertEquals(GridGrindProblemCode.IO_ERROR, fallback.problems().get(1).code());
    assertEquals(CliTransportNotice.Destination.STDOUT, notice.wroteTo());
    assertEquals(Optional.of(responseDirectory.toAbsolutePath().toString()), notice.responsePath());
    assertFalse(stderr.toString(StandardCharsets.UTF_8).contains("problems"));
  }

  @Test
  void discoveryPayloadResponseFileFailureBecomesARejectedCommandOnStdout() throws Exception {
    Path responseDirectory = Files.createTempDirectory("gridgrind-payload-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writePayload(
            "print-request-template",
            Optional.of(responseDirectory),
            stdout,
            stderr,
            "{}".getBytes(StandardCharsets.UTF_8),
            0,
            false);

    CommandError fallback = commandError(stdout.toByteArray());
    CliTransportNotice notice =
        GridGrindCliJson.readBytes(stderr.toByteArray(), CliTransportNotice.class);
    assertEquals(1, exitCode);
    assertEquals("REJECTED", fallback.status());
    assertEquals(GridGrindProblemCode.IO_ERROR, fallback.primaryProblem().code());
    assertEquals(CliTransportNotice.Destination.STDOUT, notice.wroteTo());
  }

  @Test
  void executionResponseFileFailureUsesWorkbookResultAndPreservesExecutionArtifacts()
      throws Exception {
    Path responseDirectory = Files.createTempDirectory("gridgrind-result-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    WorkbookResult.Success result = WorkbookResults.success(List.of(), List.of(), List.of());

    int exitCode =
        responseWriter.write(Optional.of(responseDirectory), stdout, stderr, result, 0, false);

    WorkbookResult.Failure fallback =
        assertInstanceOf(
            WorkbookResult.Failure.class, GridGrindJson.readWorkbookResult(stdout.toByteArray()));
    CliTransportNotice notice =
        GridGrindCliJson.readBytes(stderr.toByteArray(), CliTransportNotice.class);
    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.IO_ERROR, fallback.problem().code());
    assertEquals(result.journal(), fallback.journal());
    assertEquals(result.warnings(), fallback.warnings());
    assertEquals(result.assertions(), fallback.assertions());
    assertEquals(result.inspections(), fallback.inspections());
    assertEquals(CliTransportNotice.Destination.STDOUT, notice.wroteTo());
  }

  @Test
  void doctorResponseFileFailureIsARejectedCommandRatherThanASecondDoctorEnvelope()
      throws Exception {
    Path responseDirectory = Files.createTempDirectory("gridgrind-doctor-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    RequestDoctorReport report =
        RequestDoctorReport.invalid(Optional.empty(), List.of(), requestProblem("invalid request"));

    int exitCode =
        responseWriter.writeDoctorReport(
            Optional.of(responseDirectory), stdout, stderr, report, false);

    CommandError fallback = commandError(stdout.toByteArray());
    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.IO_ERROR, fallback.primaryProblem().code());
    assertEquals(
        CliTransportNotice.Destination.STDOUT,
        GridGrindCliJson.readBytes(stderr.toByteArray(), CliTransportNotice.class).wroteTo());
  }

  @Test
  void stdoutFallbackRemainsUsableWhenItsOptionalStderrTransportNoticeCannotBeWritten()
      throws Exception {
    Path responseDirectory = Files.createTempDirectory("gridgrind-fallback-stderr-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writePayload(
            "print-request-template",
            Optional.of(responseDirectory),
            stdout,
            failingOutputStream(),
            "{}".getBytes(StandardCharsets.UTF_8),
            0,
            false);

    assertEquals(1, exitCode);
    assertEquals(
        GridGrindProblemCode.IO_ERROR, commandError(stdout.toByteArray()).primaryProblem().code());
  }

  private static CommandError commandError() {
    return new CommandError(
        GridGrindProtocolVersion.current(),
        "execute",
        List.of(
            GridGrindProblemDetail.Problem.of(
                GridGrindProblemCode.INVALID_ARGUMENTS,
                "Unknown argument: --bogus",
                new ProblemContext.ParseArguments(CliArgument.named("--bogus")))));
  }

  private static GridGrindProblemDetail.Problem requestProblem(String message) {
    return GridGrindProblemDetail.Problem.of(
        GridGrindProblemCode.INVALID_REQUEST,
        message,
        new ProblemContext.ParseArguments(CliArgument.named("--request")));
  }

  private static OutputStream failingOutputStream() {
    return new OutputStream() {
      @Override
      public void write(int ignored) throws IOException {
        throw new IOException("test output failure");
      }
    };
  }
}
