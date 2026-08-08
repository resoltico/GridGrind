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
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import org.junit.jupiter.api.Test;

/** Last-resort failures retain the same canonical command-error core. */
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
    CliTransportNotice notice = GridGrindCliJson.readBytes(stderr.toByteArray(), CliTransportNotice.class);
    assertEquals(1, exitCode);
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, error.primaryProblem().code());
    assertEquals(CliTransportNotice.Destination.STDOUT, notice.wroteTo());
  }
}
