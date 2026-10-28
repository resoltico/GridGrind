package dev.erst.gridgrind.cli.io;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.nio.charset.StandardCharsets;
import org.junit.jupiter.api.Test;

/** Focused coverage for payload newline normalization shared across CLI output paths. */
class CliPayloadOutputTest {
  @Test
  void writeAddsOneTrailingNewlineWhenPayloadIsMissingOne() throws IOException {
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

    CliPayloadOutput.write(outputStream, "{\"status\":\"ok\"}".getBytes(StandardCharsets.UTF_8));

    assertEquals("{\"status\":\"ok\"}\n", outputStream.toString(StandardCharsets.UTF_8));
  }

  @Test
  void writePreservesExistingTrailingNewline() throws IOException {
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

    CliPayloadOutput.write(outputStream, "{\"status\":\"ok\"}\n".getBytes(StandardCharsets.UTF_8));

    assertEquals("{\"status\":\"ok\"}\n", outputStream.toString(StandardCharsets.UTF_8));
  }

  @Test
  void writeTurnsAnEmptyPayloadIntoOneEmptyLine() throws IOException {
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

    CliPayloadOutput.write(outputStream, new byte[0]);

    assertEquals("\n", outputStream.toString(StandardCharsets.UTF_8));
  }

  @Test
  void healthyPrintStreamDeliversThePayload() throws IOException {
    var bytes = new ByteArrayOutputStream();
    try (var output = new PrintStream(bytes, false, StandardCharsets.UTF_8)) {
      CliPayloadOutput.write(output, "payload".getBytes(StandardCharsets.UTF_8));
      assertEquals("payload\n", bytes.toString(StandardCharsets.UTF_8));
    }
  }

  @Test
  void closedPrintStreamCannotSilentlyReportDelivery() {
    var output = new PrintStream(new ByteArrayOutputStream(), false, StandardCharsets.UTF_8);
    output.close();
    assertThrows(
        IOException.class,
        () -> CliPayloadOutput.write(output, "payload".getBytes(StandardCharsets.UTF_8)));
  }
}
