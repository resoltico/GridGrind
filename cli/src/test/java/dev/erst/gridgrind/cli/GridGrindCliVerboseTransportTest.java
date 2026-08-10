package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.cli.discovery.CliTransportNotice;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Verifies that verbose execution preserves the CLI's one-record stderr transport boundary. */
class GridGrindCliVerboseTransportTest extends GridGrindCliTestSupport {
  @Test
  void responseFileFallbackKeepsVerboseEventsInThePrimaryPayload() throws IOException {
    Path responseDirectory = Files.createTempDirectory("gridgrind-response-dir-");
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();

    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments("--response", responseDirectory.toString()),
                new ByteArrayInputStream(
                    requestJson(
                            "{ \"type\": \"NEW\" }",
                            "{ \"type\": \"NONE\" }",
                            verboseExecutionJson(),
                            emptyFormulaEnvironmentJson(),
                            "[]")
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    WorkbookResult.Success success =
        assertInstanceOf(WorkbookResult.Success.class, response(stdout, stderr));
    String stderrPayload = stderr.toString(StandardCharsets.UTF_8);
    CliTransportNotice transportNotice =
        GridGrindCliJson.readBytes(stderr.toByteArray(), CliTransportNotice.class);

    assertEquals(1, exitCode);
    assertFalse(
        success.journal().events().isEmpty(),
        "VERBOSE events must remain in the stdout fallback payload, not on stderr");
    assertEquals(
        1,
        stderrPayload.lines().count(),
        "stderr must contain only the one structured response-file fallback notice");
    assertEquals(CliTransportNotice.Destination.STDOUT, transportNotice.wroteTo());
    assertEquals(
        java.util.Optional.of(responseDirectory.toAbsolutePath().toString()),
        transportNotice.responsePath());
  }
}
