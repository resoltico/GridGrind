package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.io.CliResponseWriter;
import dev.erst.gridgrind.cli.render.GridGrindCliJson;
import dev.erst.gridgrind.cli.render.ProductInfo;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.Optional;

/** Owns the help command output. */
final class HelpCommand {
  private HelpCommand() {}

  static int execute(
      CliCommand.Help command,
      Optional<CliOutputFormat> outputFormat,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    byte[] payload = renderHelpPayload(command, outputFormat, prettyJson);
    return responseWriter.writePayload(command.responsePath(), stdout, stderr, payload, 0);
  }

  private static byte[] renderHelpPayload(
      CliCommand.Help command, Optional<CliOutputFormat> outputFormat, boolean prettyJson)
      throws IOException {
    if (outputFormat.orElse(CliOutputFormat.TEXT) == CliOutputFormat.STRUCTURED) {
      return GridGrindCliJson.writeBytes(ProductInfo.helpReport(command.topic()), prettyJson);
    }
    return ProductInfo.helpText(command.topic(), ProductInfo.version())
        .getBytes(StandardCharsets.UTF_8);
  }
}
