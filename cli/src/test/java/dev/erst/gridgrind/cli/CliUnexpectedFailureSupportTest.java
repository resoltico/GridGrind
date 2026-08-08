package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;

import dev.erst.gridgrind.cli.discovery.CliTransportNotice;
import dev.erst.gridgrind.cli.discovery.CommandError;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.Test;

/** Last-resort failures retain the canonical command error on stdout only. */
class CliUnexpectedFailureSupportTest extends GridGrindCliTestSupport {
  @Test
  void appConvertsRunnerCrashesIntoACommandErrorOnStdout() throws IOException {
    AtomicInteger observedExitCode = new AtomicInteger(-1);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    App app =
        new App(
            () ->
                (args, stdin, out, err) -> {
                  throw new IOException("source-secret");
                },
            observedExitCode::set);

    app.run(new String[] {"--help"}, new ByteArrayInputStream(new byte[0]), stdout, stderr);

    CommandError error = commandErrorOnStdout(stdout, stderr);
    assertEquals(1, observedExitCode.get());
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, error.primaryProblem().code());
    assertFalse(new String(stdout.toByteArray(), StandardCharsets.UTF_8).contains("source-secret"));
  }

  @Test
  void responseFileFailureRecoversTheCommandErrorOnStdoutAndEmitsOneNotice() throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-unexpected-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        CliUnexpectedFailureSupport.emit(
            new String[] {"--help", "--response", responseDirectory.toString()},
            Optional.of(responseDirectory),
            false,
            stdout,
            stderr,
            new IllegalStateException("secret"));

    CommandError error = commandError(stdout.toByteArray());
    CliTransportNotice notice =
        GridGrindCliJson.readBytes(stderr.toByteArray(), CliTransportNotice.class);
    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, error.primaryProblem().code());
    assertEquals(CliTransportNotice.Destination.STDOUT, notice.wroteTo());
  }

  @Test
  void emergencyFallbackDoesNotMoveTheCommandErrorToStderrWhenStdoutIsUnavailable()
      throws Exception {
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        CliUnexpectedFailureSupport.emit(
            new String[] {"--help"},
            Optional.empty(),
            false,
            failingOutputStream(),
            stderr,
            new IllegalStateException("secret"));

    assertEquals(1, exitCode);
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
  }

  @Test
  void directFallbackRetainsStdoutWhenTheOptionalTransportNoticeCannotBeWritten() throws Exception {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    Path responsePath = Files.createTempFile("gridgrind-direct-fallback-", ".json");

    CliUnexpectedFailureSupport.directFallback(
        CommandErrors.unexpectedFailure("execute", new IllegalStateException("secret")),
        stdout,
        failingOutputStream(),
        Optional.of(responsePath),
        false);

    assertEquals(
        GridGrindProblemCode.INTERNAL_ERROR,
        commandError(stdout.toByteArray()).primaryProblem().code());
  }

  @Test
  void directFallbackWritesOnlyItsTransportNoticeWhenAResponsePathIsKnown() throws Exception {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    Path responsePath = Files.createTempFile("gridgrind-direct-fallback-notice-", ".json");
    CommandError error =
        CommandErrors.unexpectedFailure("execute", new IllegalStateException("secret"));

    CliUnexpectedFailureSupport.directFallback(
        error, stdout, stderr, Optional.of(responsePath), false);

    assertEquals(error, commandError(stdout.toByteArray()));
    assertEquals(
        CliTransportNotice.Destination.STDOUT,
        GridGrindCliJson.readBytes(stderr.toByteArray(), CliTransportNotice.class).wroteTo());
  }

  @Test
  void directFallbackWritesNoStderrPayloadWhenNoResponsePathWasRequested() throws Exception {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    CliUnexpectedFailureSupport.directFallback(
        CommandErrors.unexpectedFailure("execute", new IllegalStateException("secret")),
        stdout,
        stderr,
        Optional.empty(),
        false);

    assertEquals(
        GridGrindProblemCode.INTERNAL_ERROR,
        commandError(stdout.toByteArray()).primaryProblem().code());
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
  }

  @Test
  void directFallbackDoesNotAttemptStderrWhenStdoutIsUnavailable() throws IOException {
    AtomicInteger stdoutAttempts = new AtomicInteger();
    AtomicInteger stderrAttempts = new AtomicInteger();

    CliUnexpectedFailureSupport.directFallback(
        CommandErrors.unexpectedFailure("execute", new IllegalStateException("secret")),
        failingOutputStream(stdoutAttempts),
        failingOutputStream(stderrAttempts),
        Optional.empty(),
        false);

    assertEquals(1, stdoutAttempts.get());
    assertEquals(0, stderrAttempts.get());
  }

  private static OutputStream failingOutputStream() {
    return failingOutputStream(new AtomicInteger());
  }

  private static OutputStream failingOutputStream(AtomicInteger writeAttempts) {
    return new OutputStream() {
      @Override
      public void write(int ignored) throws IOException {
        writeAttempts.incrementAndGet();
        throw new IOException("test output failure");
      }
    };
  }
}
