package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.Optional;

/** Help, version, license, and request-template CLI surfaces. */
final class GridGrindCliIdentityCommands {
  private GridGrindCliIdentityCommands() {}

  static int help(
      CliCommand.Help command,
      Optional<CliOutputFormat> outputFormat,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    byte[] payload = renderHelpPayload(command, outputFormat, prettyJson);
    return CliCatalogPayloadSupport.writePayload(
        responseWriter, command.responsePath(), stdout, stderr, payload);
  }

  static int version(
      CliCommand.Version command,
      Optional<CliOutputFormat> outputFormat,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    byte[] payload = renderVersionPayload(outputFormat, prettyJson);
    return CliCatalogPayloadSupport.writePayload(
        responseWriter, command.responsePath(), stdout, stderr, payload);
  }

  static int license(
      CliCommand.License command,
      Optional<CliOutputFormat> outputFormat,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    byte[] payload = renderLicensePayload(outputFormat, prettyJson);
    return CliCatalogPayloadSupport.writePayload(
        responseWriter, command.responsePath(), stdout, stderr, payload);
  }

  static int requestTemplate(
      CliCommand.PrintRequestTemplate command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return CliCatalogPayloadSupport.writePayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        GridGrindJsonOutput.writeRequestBytes(
            GridGrindProtocolCatalog.requestTemplate(), prettyJson));
  }

  private static byte[] renderHelpPayload(
      CliCommand.Help command, Optional<CliOutputFormat> outputFormat, boolean prettyJson)
      throws IOException {
    if (CliCatalogPayloadSupport.effectiveTextSurfaceFormat(outputFormat)
        == CliOutputFormat.STRUCTURED) {
      return GridGrindCliJson.writeBytes(
          GridGrindCliProductInfo.helpReport(command.topic()), prettyJson);
    }
    return GridGrindCliProductInfo.helpText(command.topic(), GridGrindCliProductInfo.version())
        .getBytes(StandardCharsets.UTF_8);
  }

  private static byte[] renderVersionPayload(
      Optional<CliOutputFormat> outputFormat, boolean prettyJson) throws IOException {
    if (CliCatalogPayloadSupport.effectiveTextSurfaceFormat(outputFormat)
        == CliOutputFormat.STRUCTURED) {
      return GridGrindCliJson.writeBytes(GridGrindCliProductInfo.versionReport(), prettyJson);
    }
    String version = GridGrindCliProductInfo.version();
    String description = GridGrindCliProductInfo.description();
    return GridGrindCliProductInfo.productHeader(version, description)
        .getBytes(StandardCharsets.UTF_8);
  }

  private static byte[] renderLicensePayload(
      Optional<CliOutputFormat> outputFormat, boolean prettyJson) throws IOException {
    if (CliCatalogPayloadSupport.effectiveTextSurfaceFormat(outputFormat)
        == CliOutputFormat.STRUCTURED) {
      return GridGrindCliJson.writeBytes(GridGrindCliProductInfo.licenseReport(), prettyJson);
    }
    return GridGrindCliProductInfo.licenseText(GridGrindCli.class).getBytes(StandardCharsets.UTF_8);
  }
}
