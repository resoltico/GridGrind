package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.io.CliResponseWriter;
import dev.erst.gridgrind.cli.render.ProtocolCatalogCliJson;
import dev.erst.gridgrind.core.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.json.codec.GridGrindJsonOutput;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Optional;

/** Protocol-catalog CLI surfaces. */
final class CatalogCommand {
  private CatalogCommand() {}

  static int protocolCatalogIndex(
      CliCommand.CatalogIndex command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return responseWriter.writeRenderedPayload(
        command.responsePath(),
        stdout,
        stderr,
        output ->
            ProtocolCatalogCliJson.writeProtocolCatalogIndexReport(
                output, CliCatalogCommandSupport.protocolCatalogIndexReport(), prettyJson));
  }

  static int protocolCatalogSearch(
      CliCommand.CatalogSearch command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return responseWriter.writeRenderedPayload(
        command.responsePath(),
        stdout,
        stderr,
        output ->
            ProtocolCatalogCliJson.writeProtocolCatalogSearchReport(
                output,
                CliCatalogCommandSupport.summarizedSearchReport(command.searchQuery()),
                prettyJson));
  }

  static int protocolCatalogLookup(
      CliCommand.CatalogShow command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    java.util.List<String> matches = GridGrindProtocolCatalog.matchingLookupIds(command.lookupId());
    if (matches.size() > 1) {
      String message =
          "Ambiguous lookup id: "
              + command.lookupId()
              + ". Use one of: "
              + String.join(", ", matches);
      return responseWriter.writeCommandError(
          command.responsePath(),
          stdout,
          stderr,
          CommandErrors.invalidArguments("catalog", Optional.of("id"), message),
          prettyJson);
    }
    var lookupValue = GridGrindProtocolCatalog.lookupValueFor(command.lookupId());
    if (lookupValue.isEmpty()) {
      String message = CliCatalogCommandSupport.unknownOperationMessage(command.lookupId());
      return responseWriter.writeCommandError(
          command.responsePath(),
          stdout,
          stderr,
          CommandErrors.invalidArguments("catalog", Optional.of("id"), message),
          prettyJson);
    }
    return responseWriter.writeRenderedPayload(
        command.responsePath(),
        stdout,
        stderr,
        output ->
            GridGrindJsonOutput.writeCatalogLookupResult(
                output,
                GridGrindProtocolCatalog.catalog().protocolVersion(),
                lookupValue.get(),
                GridGrindProtocolCatalog.referencedNotesForLookupValue(lookupValue.get()),
                prettyJson));
  }
}
