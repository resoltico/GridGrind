package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.io.CliResponseWriter;
import dev.erst.gridgrind.cli.render.GridGrindCliJson;
import dev.erst.gridgrind.cli.render.ProductInfo;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.Optional;

/** Owns the version command output. */
final class VersionCommand {
  private VersionCommand() {}

  static int execute(
      CliCommand.Version command,
      Optional<CliOutputFormat> outputFormat,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    byte[] payload = renderVersionPayload(outputFormat, prettyJson);
    return responseWriter.writePayload(command.responsePath(), stdout, stderr, payload, 0);
  }

  private static byte[] renderVersionPayload(
      Optional<CliOutputFormat> outputFormat, boolean prettyJson) throws IOException {
    if (outputFormat.orElse(CliOutputFormat.TEXT) == CliOutputFormat.STRUCTURED) {
      return GridGrindCliJson.writeBytes(ProductInfo.versionReport(), prettyJson);
    }
    String version = ProductInfo.version();
    String description = ProductInfo.description();
    return ProductInfo.productHeader(version, description).getBytes(StandardCharsets.UTF_8);
  }
}
