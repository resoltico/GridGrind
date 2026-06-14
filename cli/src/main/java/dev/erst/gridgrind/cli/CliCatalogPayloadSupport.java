package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.Optional;

/** Shared payload emission support for non-executing CLI commands. */
final class CliCatalogPayloadSupport {
  private CliCatalogPayloadSupport() {}

  static int writePayload(
      CliResponseWriter responseWriter,
      String command,
      String payloadName,
      Optional<String> stdoutSuggestion,
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      byte[] payload)
      throws IOException {
    return responseWriter.writePayload(
        command, payloadName, stdoutSuggestion, responsePath, stdout, stderr, payload, 0);
  }

  static int writeRenderedPayload(
      CliResponseWriter responseWriter,
      String command,
      String payloadName,
      Optional<String> stdoutSuggestion,
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      OutputRenderer renderer)
      throws IOException {
    ByteArrayOutputStream buffer = new ByteArrayOutputStream();
    renderer.write(buffer);
    return responseWriter.writePayload(
        command,
        payloadName,
        stdoutSuggestion,
        responsePath,
        stdout,
        stderr,
        buffer.toByteArray(),
        0);
  }

  static int writeCliFailure(
      CliResponseWriter responseWriter,
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      CliFailureReport failureReport)
      throws IOException {
    return responseWriter.writeCliFailureReport(responsePath, stdout, stderr, failureReport);
  }

  static CliOutputFormat effectiveTextSurfaceFormat(Optional<CliOutputFormat> outputFormat) {
    return outputFormat.orElse(CliOutputFormat.TEXT);
  }

  /** Renders one command-specific payload into the caller-owned output buffer. */
  @FunctionalInterface
  interface OutputRenderer {
    /** Writes one command payload into the supplied output stream. */
    void write(OutputStream outputStream) throws IOException;
  }
}
