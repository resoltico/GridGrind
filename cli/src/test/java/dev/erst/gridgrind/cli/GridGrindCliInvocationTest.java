package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import org.junit.jupiter.api.Test;

/** Focused invocation-path tests for stdin discovery and execution behavior. */
class GridGrindCliInvocationTest {
  @Test
  void noArgInvocationWithEmptyStandardInputPrintsHelp() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli().run(new String[0], new ByteArrayInputStream(new byte[0]), stdout);

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
        new GridGrindCli()
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
}
