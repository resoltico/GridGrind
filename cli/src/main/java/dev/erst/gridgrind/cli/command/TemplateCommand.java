package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.io.CliResponseWriter;
import dev.erst.gridgrind.core.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.json.codec.GridGrindJsonOutput;
import java.io.IOException;
import java.io.OutputStream;

/** Owns the template command output. */
final class TemplateCommand {
  private TemplateCommand() {}

  static int request(
      CliCommand.TemplateRequest command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return responseWriter.writePayload(
        command.responsePath(),
        stdout,
        stderr,
        GridGrindJsonOutput.writeRequestBytes(
            GridGrindProtocolCatalog.requestTemplate(), prettyJson),
        0);
  }

  static int grant(
      CliCommand.TemplateGrant command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    ExecutionGrant template = ExecutionGrant.denied();
    return responseWriter.writePayload(
        command.responsePath(),
        stdout,
        stderr,
        GridGrindJsonOutput.writeGrantBytes(template, prettyJson),
        0);
  }
}
