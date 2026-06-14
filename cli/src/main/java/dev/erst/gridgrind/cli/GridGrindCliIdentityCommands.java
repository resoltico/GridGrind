package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.GridGrindCliJson;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.json.GridGrindJson;
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
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    byte[] payload = renderHelpPayload(command, outputFormat);
    return CliCatalogPayloadSupport.writePayload(
        responseWriter,
        "help",
        "help text",
        Optional.of("gridgrind --help"),
        command.responsePath(),
        stdout,
        stderr,
        payload);
  }

  static int version(
      CliCommand.Version command,
      Optional<CliOutputFormat> outputFormat,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    byte[] payload = renderVersionPayload(outputFormat);
    return CliCatalogPayloadSupport.writePayload(
        responseWriter,
        "version",
        "version output",
        Optional.of("gridgrind --version"),
        command.responsePath(),
        stdout,
        stderr,
        payload);
  }

  static int license(
      CliCommand.License command,
      Optional<CliOutputFormat> outputFormat,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    byte[] payload = renderLicensePayload(outputFormat);
    return CliCatalogPayloadSupport.writePayload(
        responseWriter,
        "license",
        "license output",
        Optional.of("gridgrind --license"),
        command.responsePath(),
        stdout,
        stderr,
        payload);
  }

  static int requestTemplate(
      CliCommand.PrintRequestTemplate command,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return CliCatalogPayloadSupport.writePayload(
        responseWriter,
        "print-request-template",
        "request template",
        Optional.of("gridgrind --print-request-template"),
        command.responsePath(),
        stdout,
        stderr,
        GridGrindJson.writeRequestBytes(GridGrindProtocolCatalog.requestTemplate()));
  }

  private static byte[] renderHelpPayload(
      CliCommand.Help command, Optional<CliOutputFormat> outputFormat) throws IOException {
    if (CliCatalogPayloadSupport.effectiveTextSurfaceFormat(outputFormat)
        == CliOutputFormat.STRUCTURED) {
      return GridGrindCliJson.writeBytes(GridGrindCliProductInfo.helpReport(command.topic()));
    }
    return GridGrindCliProductInfo.helpText(command.topic(), GridGrindCliProductInfo.version())
        .getBytes(StandardCharsets.UTF_8);
  }

  private static byte[] renderVersionPayload(Optional<CliOutputFormat> outputFormat)
      throws IOException {
    if (CliCatalogPayloadSupport.effectiveTextSurfaceFormat(outputFormat)
        == CliOutputFormat.STRUCTURED) {
      return GridGrindCliJson.writeBytes(GridGrindCliProductInfo.versionReport());
    }
    String version = GridGrindCliProductInfo.version();
    String description = GridGrindCliProductInfo.description();
    return GridGrindCliProductInfo.productHeader(version, description)
        .getBytes(StandardCharsets.UTF_8);
  }

  private static byte[] renderLicensePayload(Optional<CliOutputFormat> outputFormat)
      throws IOException {
    if (CliCatalogPayloadSupport.effectiveTextSurfaceFormat(outputFormat)
        == CliOutputFormat.STRUCTURED) {
      return GridGrindCliJson.writeBytes(GridGrindCliProductInfo.licenseReport());
    }
    return GridGrindCliProductInfo.licenseText(GridGrindCli.class).getBytes(StandardCharsets.UTF_8);
  }
}
