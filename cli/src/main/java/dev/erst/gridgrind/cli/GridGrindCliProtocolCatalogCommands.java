package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.ProtocolCatalogCliJson;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Optional;

/** Protocol-catalog CLI surfaces. */
final class GridGrindCliProtocolCatalogCommands {
  private GridGrindCliProtocolCatalogCommands() {}

  static int protocolCatalogIndex(
      CliCommand.PrintProtocolCatalogIndex command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return CliCatalogPayloadSupport.writeRenderedPayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        output ->
            ProtocolCatalogCliJson.writeProtocolCatalogIndexReport(
                output, CliCatalogCommandSupport.protocolCatalogIndexReport(), prettyJson),
        prettyJson);
  }

  static int protocolCatalogSearch(
      CliCommand.PrintProtocolCatalogSearch command,
      boolean prettyJson,
      OutputStream stdout,
      OutputStream stderr,
      CliResponseWriter responseWriter)
      throws IOException {
    return CliCatalogPayloadSupport.writeRenderedPayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        output ->
            ProtocolCatalogCliJson.writeProtocolCatalogSearchReport(
                output,
                CliCatalogCommandSupport.summarizedSearchReport(command.searchQuery()),
                prettyJson),
        prettyJson);
  }

  static int protocolCatalogLookup(
      CliCommand.PrintProtocolCatalogLookup command,
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
      return CliCatalogPayloadSupport.writeCommandError(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CommandErrors.invalidArguments(
              "print-protocol-catalog", Optional.of("--lookup"), message),
          prettyJson);
    }
    var lookupValue = GridGrindProtocolCatalog.lookupValueFor(command.lookupId());
    if (lookupValue.isEmpty()) {
      String message = CliCatalogCommandSupport.unknownOperationMessage(command.lookupId());
      return CliCatalogPayloadSupport.writeCommandError(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CommandErrors.invalidArguments(
              "print-protocol-catalog", Optional.of("--lookup"), message),
          prettyJson);
    }
    return CliCatalogPayloadSupport.writeRenderedPayload(
        responseWriter,
        command.responsePath(),
        stdout,
        stderr,
        output ->
            GridGrindJsonOutput.writeCatalogLookupResult(
                output,
                GridGrindProtocolCatalog.catalog().protocolVersion(),
                lookupValue.get(),
                GridGrindProtocolCatalog.referencedNotesForLookupValue(lookupValue.get()),
                prettyJson),
        prettyJson);
  }
}
