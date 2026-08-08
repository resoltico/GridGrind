package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CommandError;
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
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      byte[] payload,
      boolean prettyJson)
      throws IOException {
    return responseWriter.writePayload(
        command, responsePath, stdout, stderr, payload, 0, prettyJson);
  }

  static int writeRenderedPayload(
      CliResponseWriter responseWriter,
      String command,
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      OutputRenderer renderer,
      boolean prettyJson)
      throws IOException {
    ByteArrayOutputStream buffer = new ByteArrayOutputStream();
    renderer.write(buffer);
    return responseWriter.writePayload(
        command, responsePath, stdout, stderr, buffer.toByteArray(), 0, prettyJson);
  }

  static int writeCommandError(
      CliResponseWriter responseWriter,
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      CommandError commandError,
      boolean prettyJson)
      throws IOException {
    return responseWriter.writeCommandError(responsePath, stdout, stderr, commandError, prettyJson);
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
