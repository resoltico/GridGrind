package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.cli.discovery.CommandError;
import dev.erst.gridgrind.contract.dto.CliRuntimeContext;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.WorkbookResults;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies that fatal failures never fabricate an execution response contract. */
class GridGrindCliUnexpectedFailureTest extends GridGrindCliTestSupport {
  @Test
  void executionStartedErrorsProduceRejectedCommandErrorsInsteadOfFabricatedJournals()
      throws Exception {
    GridGrindCli cli =
        GridGrindCli.forTesting(
            (request, inputs, sink) -> {
              throw new AssertionError("secret execution failure");
            });
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        cli.run(
            stdinExecutionArguments(),
            new ByteArrayInputStream(
                requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                    .getBytes(StandardCharsets.UTF_8)),
            stdout,
            stderr);

    assertEquals(1, exitCode);
    CommandError failure = commandErrorOnStdout(stdout, stderr);
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.primaryProblem().code());
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR.title(), failure.primaryProblem().message());
    assertEquals(
        GridGrindProblemCode.INTERNAL_ERROR.title(),
        failure.primaryProblem().causes().getFirst().message());
    assertFalse(stdout.toString(StandardCharsets.UTF_8).contains("secret execution failure"));
    assertEquals("", stderr.toString(StandardCharsets.UTF_8));
  }

  @Test
  void preExecutionErrorsRemainRejectedCommandErrors(@TempDir Path temporaryDirectory)
      throws Exception {
    GridGrindCli cli =
        GridGrindCli.forTesting(
            (request, inputs, sink) -> {
              throw new AssertionError("executor must not run");
            },
            () -> {
              throw new AssertionError("secret pre-execution failure");
            });
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    Path executionRoot = temporaryDirectory.resolve("execution-root");

    int exitCode =
        cli.run(
            new String[] {"--doctor-request", "--execution-root", executionRoot.toString()},
            new ByteArrayInputStream(new byte[0]),
            stdout,
            stderr);

    var rejected = commandErrorOnStdout(stdout, stderr);
    assertEquals(1, exitCode);
    assertEquals("REJECTED", rejected.status());
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, rejected.primaryProblem().code());
    assertInstanceOf(CliRuntimeContext.class, rejected.primaryProblem().context());
    assertFalse(stdout.toString(StandardCharsets.UTF_8).contains("secret pre-execution failure"));
  }

  @Test
  void primaryStdoutFailureDoesNotAppendASecondDiagnostic() throws Exception {
    GridGrindCli cli =
        GridGrindCli.forTesting(
            (request, inputs, sink) -> WorkbookResults.success(List.of(), List.of(), List.of()));
    try (PartiallyFailingOutputStream stdout = new PartiallyFailingOutputStream();
        ByteArrayOutputStream stderr = new ByteArrayOutputStream()) {
      int exitCode =
          cli.run(
              stdinExecutionArguments(),
              new ByteArrayInputStream(
                  requestJson("{ \"type\": \"NEW\" }", "{ \"type\": \"NONE\" }", "[]")
                      .getBytes(StandardCharsets.UTF_8)),
              stdout,
              stderr);

      assertEquals(1, exitCode);
      assertEquals(1, stdout.writeAttempts());
      assertEquals(16, stdout.writtenByteCount());
      assertEquals("", stderr.toString(StandardCharsets.UTF_8));
    }
  }

  /** Captures one partial write and rejects every subsequent primary-output attempt. */
  private static final class PartiallyFailingOutputStream extends OutputStream {
    private final ByteArrayOutputStream captured = new ByteArrayOutputStream();
    private int writeAttempts;

    @Override
    public void write(int value) throws IOException {
      write(new byte[] {(byte) value}, 0, 1);
    }

    @Override
    public void write(byte[] bytes, int offset, int length) throws IOException {
      writeAttempts++;
      if (writeAttempts == 1) {
        captured.write(bytes, offset, Math.min(length, 16));
      }
      throw new IOException("simulated primary output failure");
    }

    int writeAttempts() {
      return writeAttempts;
    }

    int writtenByteCount() {
      return captured.size();
    }
  }
}
