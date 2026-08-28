package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.CliTransportNotice;
import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.dto.ExecutionProgressEvent;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Verifies that verbose progress and transport notices share the JSONL stderr channel. */
class GridGrindCliVerboseTransportTest extends GridGrindCliTestSupport {
  @Test
  void responseFileFallbackKeepsVerboseProgressAsCompactJsonLines() throws IOException {
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

    String stderrPayload = stderr.toString(StandardCharsets.UTF_8);
    java.util.List<String> stderrLines = stderrPayload.lines().toList();
    CliTransportNotice transportNotice =
        GridGrindCliJson.readBytes(
            stderrLines.getLast().getBytes(StandardCharsets.UTF_8), CliTransportNotice.class);

    assertEquals(1, exitCode);
    assertEquals(1, stderrLines.size());
    assertEquals("", stdout.toString(StandardCharsets.UTF_8));
    assertEquals(CliTransportNotice.Destination.NOT_DELIVERED, transportNotice.wroteTo());
    assertEquals(CliTransportNotice.Reason.RESPONSE_PATH_DIRECTORY, transportNotice.reason());
    assertEquals(
        java.util.Optional.of(responseDirectory.toAbsolutePath().toString()),
        transportNotice.responsePath());
  }

  @Test
  void verboseProgressStaysCompactUnderPrettyAndDoesNotEchoSecrets() throws IOException {
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    ByteArrayOutputStream stderr = new ByteArrayOutputStream();
    String sentinel = "progress-secret-sentinel";

    int exitCode =
        new GridGrindCli()
            .run(
                stdinExecutionArguments("--pretty"),
                new ByteArrayInputStream(
                    requestJson(
                            """
                            {
                              "type": "EXISTING",
                              "path": "missing.xlsx",
                              "security": { "password": "%s" }
                            }
                            """
                                .formatted(sentinel),
                            "{ \"type\": \"NONE\" }",
                            verboseExecutionJson(),
                            emptyFormulaEnvironmentJson(),
                            "[]")
                        .getBytes(StandardCharsets.UTF_8)),
                stdout,
                stderr);

    java.util.List<String> progressLines = stderr.toString(StandardCharsets.UTF_8).lines().toList();

    assertEquals(1, exitCode);
    assertTrue(stdout.toString(StandardCharsets.UTF_8).contains("\n  \"status\""));
    assertFalse(progressLines.isEmpty());
    assertTrue(
        progressLines.stream().allMatch(line -> line.startsWith("{") && !line.contains("\n")));
    assertTrue(
        progressLines.stream()
            .map(
                line -> {
                  try {
                    return GridGrindCliJson.readBytes(
                        line.getBytes(StandardCharsets.UTF_8), ExecutionProgressEvent.class);
                  } catch (IOException exception) {
                    throw new java.io.UncheckedIOException(exception);
                  }
                })
            .allMatch(event -> event.status() != null && event.category() != null));
    assertFalse(stderr.toString(StandardCharsets.UTF_8).contains(sentinel));
    assertFalse(stderr.toString(StandardCharsets.UTF_8).contains("\"detail\""));
  }

  @Test
  void progressSinkFailsClosedWhenStderrCannotAcceptJsonl() throws IOException {
    IOException failure = new IOException("stderr unavailable");
    try (OutputStream unavailable =
        new OutputStream() {
          @Override
          public void write(int ignored) throws IOException {
            throw failure;
          }
        }) {
      CliPrimaryOutputException exception =
          assertThrows(
              CliPrimaryOutputException.class,
              () ->
                  new CliProgressJsonlSink(unavailable)
                      .emit(
                          ExecutionProgressEvent.started(
                              "2026-08-20T17:00:00Z",
                              ExecutionProgressEvent.Category.PLAN,
                              java.util.Optional.empty(),
                              java.util.Optional.empty())));

      assertSame(failure, exception.getCause());
    }
  }
}
