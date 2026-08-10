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
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResults;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.RequestDiagnosticRedactor;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
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
  void rejectedCommandResponseFileFailurePreservesTheOriginalCommandError() throws Exception {
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
    assertEquals(commandError(), fallback);
    assertEquals(CliTransportNotice.Destination.STDOUT, notice.wroteTo());
    assertEquals(Optional.of(responseDirectory.toAbsolutePath().toString()), notice.responsePath());
    assertFalse(stderr.toString(StandardCharsets.UTF_8).contains("problems"));
  }

  @Test
  void discoveryPayloadResponseFileFailurePreservesTheOriginalPayload() throws Exception {
    Path responseDirectory = Files.createTempDirectory("gridgrind-payload-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        responseWriter.writePayload(
            Optional.of(responseDirectory),
            stdout,
            stderr,
            "{}".getBytes(StandardCharsets.UTF_8),
            0);

    CliTransportNotice notice =
        GridGrindCliJson.readBytes(stderr.toByteArray(), CliTransportNotice.class);
    assertEquals(1, exitCode);
    assertEquals("{}\n", stdout.toString(StandardCharsets.UTF_8));
    assertEquals(CliTransportNotice.Destination.STDOUT, notice.wroteTo());
  }

  @Test
  void executionResponseFileFailurePreservesTheOriginalSuccessfulWorkbookResult() throws Exception {
    Path responseDirectory = Files.createTempDirectory("gridgrind-result-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    WorkbookResult.Success result = WorkbookResults.success(List.of(), List.of(), List.of());

    int exitCode =
        responseWriter.write(Optional.of(responseDirectory), stdout, stderr, result, 0, false);

    WorkbookResult.Success fallback =
        assertInstanceOf(
            WorkbookResult.Success.class, GridGrindJson.readWorkbookResult(stdout.toByteArray()));
    CliTransportNotice notice =
        GridGrindCliJson.readBytes(stderr.toByteArray(), CliTransportNotice.class);
    assertEquals(1, exitCode);
    assertEquals(result, fallback);
    assertEquals(CliTransportNotice.Destination.STDOUT, notice.wroteTo());
  }

  @Test
  void doctorResponseFileFailurePreservesTheOriginalDoctorReport() throws Exception {
    Path responseDirectory = Files.createTempDirectory("gridgrind-doctor-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    RequestDoctorReport report =
        RequestDoctorReport.invalid(Optional.empty(), List.of(), requestProblem("invalid request"));

    int exitCode =
        responseWriter.writeDoctorReport(
            Optional.of(responseDirectory), stdout, stderr, report, false);

    RequestDoctorReport fallback = GridGrindJson.readRequestDoctorReport(stdout.toByteArray());
    assertEquals(1, exitCode);
    assertEquals(report, fallback);
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
            Optional.of(responseDirectory),
            stdout,
            failingOutputStream(),
            "{}".getBytes(StandardCharsets.UTF_8),
            0);

    assertEquals(1, exitCode);
    assertEquals("{}\n", stdout.toString(StandardCharsets.UTF_8));
  }

  @Test
  void redactsEveryDeclaredSecretAcrossFilesStdoutAndFallbacks() throws Exception {
    RequestDiagnosticRedactor redactor = allSecretsRedactor();

    for (SecretOwner owner : secretOwners()) {
      assertSecretOwnerIsRedactedAcrossEveryPrimaryTransport(redactor, owner);
    }
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

  private static GridGrindProblemDetail.Problem secretProblem(SecretOwner owner) {
    GridGrindProblemCode code = GridGrindProblemCode.INVALID_REQUEST;
    return new GridGrindProblemDetail.Problem(
        code,
        code.category(),
        code.recovery(),
        code.title(),
        "Request secret was " + owner.value(),
        "Replace secret " + owner.value() + " before retrying.",
        new ProblemContext.ReadRequest(
            RequestInput.standardInput(), JsonLocation.pathOnly(owner.jsonPath())),
        Optional.empty(),
        List.of(
            new GridGrindProblemDetail.ProblemCause(
                code, "Cause: " + owner.value(), "READ_REQUEST")));
  }

  private void assertSecretOwnerIsRedactedAcrossEveryPrimaryTransport(
      RequestDiagnosticRedactor redactor, SecretOwner owner) throws IOException {
    GridGrindProblemDetail.Problem problem = secretProblem(owner);
    CommandError commandError =
        new CommandError(GridGrindProtocolVersion.current(), "execute", List.of(problem));
    RequestDoctorReport doctorReport =
        RequestDoctorReport.invalid(Optional.empty(), List.of(), problem);
    WorkbookResult.Failure executionResult = WorkbookResults.failure(problem);

    ByteArrayOutputStream commandStdout = new ByteArrayOutputStream();
    ByteArrayOutputStream commandStderr = new ByteArrayOutputStream();
    assertEquals(
        2,
        responseWriter.writeCommandError(
            Optional.empty(), commandStdout, commandStderr, commandError, redactor, false));
    assertSecretRedacted(owner, commandStdout.toByteArray(), commandStderr.toByteArray());
    assertSecretProblemRedacted(commandError(commandStdout.toByteArray()).primaryProblem());

    Path doctorResponse =
        Files.createTempDirectory("gridgrind-secret-doctor-").resolve("report.json");
    ByteArrayOutputStream doctorStdout = new ByteArrayOutputStream();
    ByteArrayOutputStream doctorStderr = new ByteArrayOutputStream();
    assertEquals(
        1,
        responseWriter.writeDoctorReport(
            Optional.of(doctorResponse),
            doctorStdout,
            doctorStderr,
            doctorReport,
            redactor,
            false));
    byte[] doctorBytes = Files.readAllBytes(doctorResponse);
    assertSecretRedacted(
        owner, doctorBytes, doctorStdout.toByteArray(), doctorStderr.toByteArray());
    assertSecretProblemRedacted(
        GridGrindJson.readRequestDoctorReport(doctorBytes).primaryProblem().orElseThrow());

    Path executionResponse =
        Files.createTempDirectory("gridgrind-secret-execution-").resolve("response.json");
    ByteArrayOutputStream executionStdout = new ByteArrayOutputStream();
    ByteArrayOutputStream executionStderr = new ByteArrayOutputStream();
    assertEquals(
        1,
        responseWriter.write(
            Optional.of(executionResponse),
            executionStdout,
            executionStderr,
            executionResult,
            1,
            redactor,
            false));
    byte[] executionBytes = Files.readAllBytes(executionResponse);
    assertSecretRedacted(
        owner, executionBytes, executionStdout.toByteArray(), executionStderr.toByteArray());
    assertSecretProblemRedacted(
        assertInstanceOf(
                WorkbookResult.Failure.class, GridGrindJson.readWorkbookResult(executionBytes))
            .problem());

    ByteArrayOutputStream commandFallbackStdout = new ByteArrayOutputStream();
    ByteArrayOutputStream commandFallbackStderr = new ByteArrayOutputStream();
    assertEquals(
        1,
        responseWriter.writeCommandError(
            Optional.of(Files.createTempDirectory("gridgrind-secret-command-fallback-")),
            commandFallbackStdout,
            commandFallbackStderr,
            commandError,
            redactor,
            false));
    assertSecretRedacted(
        owner, commandFallbackStdout.toByteArray(), commandFallbackStderr.toByteArray());
    assertSecretProblemRedacted(commandError(commandFallbackStdout.toByteArray()).primaryProblem());

    ByteArrayOutputStream doctorFallbackStdout = new ByteArrayOutputStream();
    ByteArrayOutputStream doctorFallbackStderr = new ByteArrayOutputStream();
    assertEquals(
        1,
        responseWriter.writeDoctorReport(
            Optional.of(Files.createTempDirectory("gridgrind-secret-doctor-fallback-")),
            doctorFallbackStdout,
            doctorFallbackStderr,
            doctorReport,
            redactor,
            false));
    assertSecretRedacted(
        owner, doctorFallbackStdout.toByteArray(), doctorFallbackStderr.toByteArray());
    assertSecretProblemRedacted(
        GridGrindJson.readRequestDoctorReport(doctorFallbackStdout.toByteArray())
            .primaryProblem()
            .orElseThrow());

    ByteArrayOutputStream executionFallbackStdout = new ByteArrayOutputStream();
    ByteArrayOutputStream executionFallbackStderr = new ByteArrayOutputStream();
    assertEquals(
        1,
        responseWriter.write(
            Optional.of(Files.createTempDirectory("gridgrind-secret-execution-fallback-")),
            executionFallbackStdout,
            executionFallbackStderr,
            executionResult,
            1,
            redactor,
            false));
    assertSecretRedacted(
        owner, executionFallbackStdout.toByteArray(), executionFallbackStderr.toByteArray());
    assertSecretProblemRedacted(
        assertInstanceOf(
                WorkbookResult.Failure.class,
                GridGrindJson.readWorkbookResult(executionFallbackStdout.toByteArray()))
            .problem());
  }

  private static RequestDiagnosticRedactor allSecretsRedactor() throws IOException {
    return GridGrindJson.analyzeRequest(
            """
            {
              "protocolVersion": "V2",
              "source": {
                "type": "EXISTING",
                "path": "source.xlsx",
                "security": { "password": "source-secret" }
              },
              "persistence": {
                "type": "SAVE_AS",
                "path": "secured.xlsx",
                "ifExists": "REJECT",
                "security": {
                  "encryption": { "password": "persistence-secret" },
                  "signature": {
                    "pkcs12Path": "keys/signing.p12",
                    "keystorePassword": "keystore-secret",
                    "keyPassword": "key-secret"
                  }
                }
              },
              "steps": []
            }
            """
                .getBytes(StandardCharsets.UTF_8))
        .diagnosticRedactor();
  }

  private static List<SecretOwner> secretOwners() {
    return List.of(
        new SecretOwner("source.security.password", "source-secret"),
        new SecretOwner("persistence.security.encryption.password", "persistence-secret"),
        new SecretOwner("persistence.security.signature.keystorePassword", "keystore-secret"),
        new SecretOwner("persistence.security.signature.keyPassword", "key-secret"));
  }

  private static void assertSecretRedacted(SecretOwner owner, byte[]... payloads) {
    assertFalse(
        Arrays.stream(payloads)
            .map(payload -> new String(payload, StandardCharsets.UTF_8))
            .anyMatch(payload -> payload.contains(owner.value())));
  }

  private static void assertSecretProblemRedacted(GridGrindProblemDetail.Problem problem) {
    assertEquals("[REDACTED]", problem.message());
    assertEquals("[REDACTED]", problem.resolution());
    assertEquals("[REDACTED]", problem.causes().getFirst().message());
  }

  private static OutputStream failingOutputStream() {
    return new OutputStream() {
      @Override
      public void write(int ignored) throws IOException {
        throw new IOException("test output failure");
      }
    };
  }

  private record SecretOwner(String jsonPath, String value) {}
}
