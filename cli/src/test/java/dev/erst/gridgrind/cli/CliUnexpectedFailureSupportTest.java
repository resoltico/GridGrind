package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
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

/** Direct coverage for last-resort CLI failure emission and entry-point catch wiring. */
class CliUnexpectedFailureSupportTest extends GridGrindCliTestSupport {
  @Test
  void appConvertsRunnerCrashesIntoStructuredFailures() throws IOException {
    AtomicInteger observedExitCode = new AtomicInteger(-1);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    App app =
        new App(
            () ->
                (args, stdin, out, err) -> {
                  throw new IOException("runner exploded");
                },
            observedExitCode::set);

    app.run(new String[] {"--help"}, new ByteArrayInputStream(new byte[0]), stdout, stderr);

    CliFailureReport failure = cliFailureOnStderr(stdout, stderr);
    assertEquals(1, observedExitCode.get());
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.code());
    assertEquals("runner exploded", failure.message());
  }

  @Test
  void gridGrindCliConvertsTransportWriteCrashesIntoStructuredFailures() throws IOException {
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                new String[] {"--help"},
                new ByteArrayInputStream(new byte[0]),
                new AlwaysFailingOutputStream("help stdout exploded"),
                stderr);

    CliFailureReport failure = cliFailure(stderr.toByteArray());
    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.code());
    assertEquals("help stdout exploded", failure.message());
  }

  @Test
  void emitFallsBackToStructuredStdoutWhenResponseFileFallbackWriteFailsOnce() throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-cli-unexpected-response-dir-");
    try (FailOnceThenCaptureOutputStream stdout = new FailOnceThenCaptureOutputStream()) {
      ByteArrayOutputStream stderr = new ByteArrayOutputStream();

      int exitCode =
          CliUnexpectedFailureSupport.emit(
              new String[] {"--help", "--response", responseDirectory.toString()},
              Optional.of(responseDirectory),
              stdout,
              stderr,
              new IllegalStateException("boom"));

      CliFailureReport failure = cliFailure(stdout.toByteArray());
      assertEquals(1, exitCode);
      assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.code());
      assertTrue(
          stderr
              .toString(StandardCharsets.UTF_8)
              .contains("Could not write response file " + responseDirectory.toAbsolutePath()));
    }
  }

  @Test
  void emitFallsBackToStructuredStderrWhenTheInitialFailureWriteFailsOnce() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    try (FailOnceThenCaptureOutputStream stderr = new FailOnceThenCaptureOutputStream()) {
      int exitCode =
          CliUnexpectedFailureSupport.emit(
              new String[] {"--help"},
              Optional.empty(),
              stdout,
              stderr,
              new IllegalStateException("stderr boom"));

      CliFailureReport failure = cliFailure(stderr.toByteArray());
      assertEquals(1, exitCode);
      assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.code());
      assertEquals("stderr boom", failure.message());
    }
  }

  @Test
  void emitFallsBackToHumanReadableMessageWhenNoStructuredChannelCanRecover() throws IOException {
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    Path responseDirectory = Files.createTempDirectory("gridgrind-cli-bad-dir-");

    int exitCode =
        CliUnexpectedFailureSupport.emit(
            new String[] {"--help", "--response", responseDirectory.toString()},
            Optional.of(responseDirectory),
            new AlwaysFailingOutputStream("stdout exploded"),
            stderr,
            new IllegalStateException("boom"));

    assertEquals(1, exitCode);
    assertTrue(
        stderr
            .toString(StandardCharsets.UTF_8)
            .contains("GridGrind failed before it could emit a structured error payload."));
  }

  @Test
  void emitReturnsCleanlyWhenEvenTheHumanReadableFallbackCannotBeWritten() {
    int exitCode =
        CliUnexpectedFailureSupport.emit(
            new String[] {"--help"},
            Optional.empty(),
            new ByteArrayOutputStream(),
            new AlwaysFailingOutputStream("stderr exploded"),
            new IllegalStateException("boom"));

    assertEquals(1, exitCode);
  }

  /** Output stream probe that always fails every write so the fallback path can be exercised. */
  private static final class AlwaysFailingOutputStream extends OutputStream {
    private final String message;

    private AlwaysFailingOutputStream(String message) {
      this.message = message;
    }

    @Override
    public void write(int value) throws IOException {
      throw new IOException(message);
    }

    @Override
    public void write(byte[] buffer, int offset, int length) throws IOException {
      throw new IOException(message);
    }
  }

  /** Output stream probe that fails the first write, then captures subsequent fallback bytes. */
  private static final class FailOnceThenCaptureOutputStream extends OutputStream {
    private final ByteArrayOutputStream delegate = new ByteArrayOutputStream();
    private boolean failed;

    @Override
    public void write(int value) throws IOException {
      if (!failed) {
        failed = true;
        throw new IOException("first write exploded");
      }
      delegate.write(value);
    }

    @Override
    public void write(byte[] buffer, int offset, int length) throws IOException {
      if (!failed) {
        failed = true;
        throw new IOException("first write exploded");
      }
      delegate.write(buffer, offset, length);
    }

    byte[] toByteArray() {
      return delegate.toByteArray();
    }
  }
}
