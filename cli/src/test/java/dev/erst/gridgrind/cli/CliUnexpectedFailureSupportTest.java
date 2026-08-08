package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.CommandError;
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

/** Direct coverage for last-resort CLI diagnostic emission and entry-point catch wiring. */
class CliUnexpectedFailureSupportTest extends GridGrindCliTestSupport {
  @Test
  void appConvertsRunnerCrashesIntoStructuredDiagnostics() throws IOException {
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

    CommandError failure = commandErrorOnStdout(stdout, stderr);
    assertEquals(1, observedExitCode.get());
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.primaryProblem().code());
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR.title(), failure.primaryProblem().message());
    assertFalse(stderr.toString(StandardCharsets.UTF_8).contains("source-secret"));
  }

  @Test
  void gridGrindCliConvertsTransportWriteCrashesIntoStructuredDiagnostics() throws IOException {
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--help"},
                new ByteArrayInputStream(new byte[0]),
                FailingOutputStream.checked("help stdout exploded"),
                stderr);

    CommandError failure = commandError(stderr.toByteArray());
    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.primaryProblem().code());
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR.title(), failure.primaryProblem().message());
  }

  @Test
  void emitFallsBackToStructuredStdoutWhenResponseFileFallbackWriteFailsOnce() throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-cli-unexpected-response-dir-");
    try (RecoveringOutputStream stdout = RecoveringOutputStream.checked(1)) {
      ByteArrayOutputStream stderr = new ByteArrayOutputStream();

      int exitCode =
          CliUnexpectedFailureSupport.emit(
              new String[] {"--help", "--response", responseDirectory.toString()},
              Optional.of(responseDirectory),
              false,
              stdout,
              stderr,
              new IllegalStateException("boom"));

      CommandError failure = commandError(stdout.toByteArray());
      CommandError stderrDiagnostic = commandError(stderr.toByteArray());
      assertEquals(1, exitCode);
      assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.primaryProblem().code());
      assertEquals(stderrDiagnostic, failure);
      assertEquals(Optional.of("STDOUT"), wroteTo(stderrDiagnostic));
    }
  }

  @Test
  void emitFallsBackToStructuredStderrWhenStdoutCannotCarryTheFailurePayload() throws IOException {
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    try (FailingOutputStream stdout = FailingOutputStream.checked("stdout exploded")) {
      int exitCode =
          CliUnexpectedFailureSupport.emit(
              new String[] {"--help"},
              Optional.empty(),
              false,
              stdout,
              stderr,
              new IllegalStateException("boom"));

      CommandError failure = commandError(stderr.toByteArray());
      assertEquals(1, exitCode);
      assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.primaryProblem().code());
      assertEquals(GridGrindProblemCode.INTERNAL_ERROR.title(), failure.primaryProblem().message());
    }
  }

  @Test
  void emitRecoversToStructuredStdoutWhenTheInitialResponsePathMirrorCrashesNonIo()
      throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-cli-unexpected-runtime-", ".json");
    Files.deleteIfExists(responsePath);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        CliUnexpectedFailureSupport.emit(
            new String[] {"--help", "--response", responsePath.toString()},
            Optional.of(responsePath),
            false,
            stdout,
            FailingOutputStream.runtime("stderr exploded"),
            new IllegalStateException("boom"));

    CommandError stdoutDiagnostic = commandError(stdout.toByteArray());
    CommandError fileDiagnostic = commandError(Files.readAllBytes(responsePath));
    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, stdoutDiagnostic.primaryProblem().code());
    assertEquals(Optional.of("STDOUT"), wroteTo(stdoutDiagnostic));
    assertEquals(Optional.of("FILE"), wroteTo(fileDiagnostic));
  }

  @Test
  void
      emitRecoversToStructuredStderrWhenTheInitialResponsePathMirrorCrashesNonIoAndStdoutAlsoFails()
          throws IOException {
    Path responsePath = Files.createTempFile("gridgrind-cli-unexpected-runtime-stderr-", ".json");
    Files.deleteIfExists(responsePath);
    try (RecoveringOutputStream stderr = RecoveringOutputStream.runtime(1)) {
      int exitCode =
          CliUnexpectedFailureSupport.emit(
              new String[] {"--help", "--response", responsePath.toString()},
              Optional.of(responsePath),
              false,
              FailingOutputStream.checked("stdout exploded"),
              stderr,
              new IllegalStateException("boom"));

      CommandError stderrDiagnostic = commandError(stderr.toByteArray());
      CommandError fileDiagnostic = commandError(Files.readAllBytes(responsePath));
      assertEquals(1, exitCode);
      assertEquals(GridGrindProblemCode.INTERNAL_ERROR, stderrDiagnostic.primaryProblem().code());
      assertEquals(Optional.empty(), stderrDiagnostic.transport());
      assertEquals(Optional.of("FILE"), wroteTo(fileDiagnostic));
    }
  }

  @Test
  void emitDoesNotAppendOneSecondStderrDiagnosticWhenResponseFallbackAlreadyRecoveredStderr()
      throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-cli-unexpected-stderr-dir-");
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        CliUnexpectedFailureSupport.emit(
            new String[] {"--help", "--response", responseDirectory.toString()},
            Optional.of(responseDirectory),
            false,
            FailingOutputStream.checked("stdout exploded"),
            stderr,
            new IllegalStateException("boom"));

    CommandError failure = commandError(stderr.toByteArray());
    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.primaryProblem().code());
    assertEquals(Optional.of("STDOUT"), wroteTo(failure));
  }

  @Test
  void emitRecoversToOneStructuredStderrDiagnosticAfterTheInitialResponseFallbackMirrorAlsoFails()
      throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-cli-unexpected-suppressed-dir-");
    try (RecoveringOutputStream stderr = RecoveringOutputStream.checked(1)) {
      int exitCode =
          CliUnexpectedFailureSupport.emit(
              new String[] {"--help", "--response", responseDirectory.toString()},
              Optional.of(responseDirectory),
              false,
              FailingOutputStream.checked("stdout exploded"),
              stderr,
              new IllegalStateException("boom"));

      CommandError failure = commandError(stderr.toByteArray());
      assertEquals(1, exitCode);
      assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.primaryProblem().code());
      assertEquals(Optional.empty(), failure.transport());
    }
  }

  @Test
  void emitRetriesStructuredFailureOnStderrAfterThePrimaryStderrWriteAndStdoutFallbackBothFail()
      throws IOException {
    try (RecoveringOutputStream stderr = RecoveringOutputStream.checked(1)) {
      int exitCode =
          CliUnexpectedFailureSupport.emit(
              new String[] {"--help"},
              Optional.empty(),
              false,
              FailingOutputStream.checked("stdout exploded"),
              stderr,
              new IllegalStateException("boom"));

      CommandError failure = commandError(stderr.toByteArray());
      assertEquals(1, exitCode);
      assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.primaryProblem().code());
      assertEquals(GridGrindProblemCode.INTERNAL_ERROR.title(), failure.primaryProblem().message());
    }
  }

  @Test
  void emitFallsBackToHumanReadableMessageWhenNoStructuredChannelCanRecover() throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-cli-bad-dir-");
    try (RecoveringOutputStream stderr = RecoveringOutputStream.checked(2)) {
      int exitCode =
          CliUnexpectedFailureSupport.emit(
              new String[] {"--help", "--response", responseDirectory.toString()},
              Optional.of(responseDirectory),
              false,
              FailingOutputStream.checked("stdout exploded"),
              stderr,
              new IllegalStateException("boom"));

      assertEquals(1, exitCode);
      assertTrue(
          stderr
              .toString(StandardCharsets.UTF_8)
              .contains(
                  "GridGrind failed before it could emit a structured error payload to the response fallback channels."));
    }
  }

  @Test
  void emitReturnsCleanlyWhenEvenTheHumanReadableFallbackCannotBeWritten() {
    int exitCode =
        CliUnexpectedFailureSupport.emit(
            new String[] {"--help"},
            Optional.empty(),
            false,
            FailingOutputStream.checked("stdout exploded"),
            FailingOutputStream.checked("stderr exploded"),
            new IllegalStateException("boom"));

    assertEquals(1, exitCode);
  }

  /** Output stream probe that always fails every write so the fallback path can be exercised. */
  private static final class FailingOutputStream extends OutputStream {
    private final String message;
    private final boolean checkedFailure;

    private FailingOutputStream(String message, boolean checkedFailure) {
      this.message = message;
      this.checkedFailure = checkedFailure;
    }

    private static FailingOutputStream checked(String message) {
      return new FailingOutputStream(message, true);
    }

    private static FailingOutputStream runtime(String message) {
      return new FailingOutputStream(message, false);
    }

    private void fail() throws IOException {
      if (checkedFailure) {
        throw new IOException(message);
      }
      throw new IllegalStateException(message);
    }

    @Override
    public void write(int value) throws IOException {
      fail();
    }

    @Override
    public void write(byte[] buffer, int offset, int length) throws IOException {
      fail();
    }

    @Override
    public void close() throws IOException {
      // Preserve no-op close so try-with-resources can use the probe cleanly.
    }
  }

  /** Output stream probe that fails a configured number of times, then captures later bytes. */
  private static final class RecoveringOutputStream extends OutputStream {
    private final ByteArrayOutputStream delegate = new ByteArrayOutputStream();
    private final boolean checkedFailure;
    private int remainingFailures;

    private RecoveringOutputStream(int remainingFailures, boolean checkedFailure) {
      this.remainingFailures = remainingFailures;
      this.checkedFailure = checkedFailure;
    }

    private static RecoveringOutputStream checked(int failures) {
      return new RecoveringOutputStream(failures, true);
    }

    private static RecoveringOutputStream runtime(int failures) {
      return new RecoveringOutputStream(failures, false);
    }

    private void maybeFail() throws IOException {
      if (remainingFailures <= 0) {
        return;
      }
      remainingFailures--;
      if (checkedFailure) {
        throw new IOException("write exploded");
      }
      throw new IllegalStateException("first runtime write exploded");
    }

    @Override
    public void write(int value) throws IOException {
      maybeFail();
      delegate.write(value);
    }

    @Override
    public void write(byte[] buffer, int offset, int length) throws IOException {
      maybeFail();
      delegate.write(buffer, offset, length);
    }

    @Override
    public void close() throws IOException {
      delegate.close();
    }

    byte[] toByteArray() {
      return delegate.toByteArray();
    }

    String toString(java.nio.charset.Charset charset) {
      return delegate.toString(charset);
    }
  }

  private static Optional<String> wroteTo(CommandError diagnostic) {
    return diagnostic
        .transport()
        .map(
            transport ->
                switch (transport) {
                  case dev.erst.gridgrind.cli.discovery.CliTransport.StandardOutput _ -> "STDOUT";
                  case dev.erst.gridgrind.cli.discovery.CliTransport.ResponseFile _ -> "FILE";
                });
  }
}
