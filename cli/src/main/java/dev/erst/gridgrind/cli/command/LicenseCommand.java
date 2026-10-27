package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.io.CliResponseWriter;
import dev.erst.gridgrind.cli.render.GridGrindCliJson;
import dev.erst.gridgrind.cli.render.ProductInfo;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.Optional;

/** Owns the license command output. */
final class LicenseCommand {
  private LicenseCommand() {}

  static int execute(
      CliCommand.License command,
      Optional<CliOutputFormat> outputFormat,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    byte[] payload = renderLicensePayload(outputFormat, prettyJson);
    return responseWriter.writePayload(command.responsePath(), stdout, stderr, payload, 0);
  }

  private static byte[] renderLicensePayload(
      Optional<CliOutputFormat> outputFormat, boolean prettyJson) throws IOException {
    if (outputFormat.orElse(CliOutputFormat.TEXT) == CliOutputFormat.STRUCTURED) {
      return GridGrindCliJson.writeBytes(ProductInfo.licenseReport(), prettyJson);
    }
    return ProductInfo.licenseText(GridGrindCli.class).getBytes(StandardCharsets.UTF_8);
  }
}
