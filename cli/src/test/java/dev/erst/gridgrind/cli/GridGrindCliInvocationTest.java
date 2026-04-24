package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Focused invocation-path tests for stdin discovery and execution behavior. */
class GridGrindCliInvocationTest {
  @Test
  void noArgInvocationWithEmptyStandardInputPrintsHelp() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        nonInteractiveCli().run(new String[0], new ByteArrayInputStream(new byte[0]), stdout);

    assertEquals(0, exitCode);
    String help = stdout.toString(StandardCharsets.UTF_8);
    assertTrue(help.contains("Usage:"));
    assertTrue(help.contains("--doctor-request"));
    assertTrue(help.contains("--print-task-catalog"));
    assertTrue(help.contains("--print-task-plan <id>"));
    assertTrue(help.contains("--print-goal-plan <goal>"));
    assertTrue(help.contains("--print-protocol-catalog"));
    assertTrue(help.contains("Minimal Valid Request:"));
  }

  @Test
  void noArgInvocationStillExecutesWhenStandardInputContainsARequest() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        nonInteractiveCli()
            .run(
                new String[0],
                new ByteArrayInputStream(
                    """
                    {
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)),
                stdout);

    assertEquals(0, exitCode);
    assertInstanceOf(
        GridGrindResponse.Success.class, GridGrindJson.readResponse(stdout.toByteArray()));
  }

  @Test
  void noArgInvocationWithInteractiveStandardInputPrintsHelpWithoutReadingInput()
      throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    try (InputStream blockingStdin =
        new InputStream() {
          @Override
          public int read() throws IOException {
            throw new AssertionError("interactive no-arg help must not read stdin");
          }

          @Override
          public int read(byte[] b, int off, int len) throws IOException {
            throw new AssertionError("interactive no-arg help must not read stdin");
          }
        }) {
      int exitCode = interactiveCli().run(new String[0], blockingStdin, stdout);

      assertEquals(0, exitCode);
      String help = stdout.toString(StandardCharsets.UTF_8);
      assertTrue(help.contains("Usage:"));
      assertTrue(help.contains("--doctor-request"));
      assertTrue(help.contains("--print-protocol-catalog"));
    }
  }

  private static GridGrindCli nonInteractiveCli() {
    return new GridGrindCli(
        (ignoredRequest, ignoredBindings, ignoredSink) ->
            GridGrindResponse.success(null, null, List.of(), List.of(), List.of()),
        new CliRequestReader(),
        new CliResponseWriter(),
        new CliJournalWriter(),
        () -> false);
  }

  private static GridGrindCli interactiveCli() {
    return new GridGrindCli(
        (ignoredRequest, ignoredBindings, ignoredSink) -> {
          throw new AssertionError("interactive no-arg help must not execute a request");
        },
        new CliRequestReader(),
        new CliResponseWriter(),
        new CliJournalWriter(),
        () -> true);
  }
}
