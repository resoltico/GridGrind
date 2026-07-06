package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.ProtocolCatalogCliJson;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
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
        "print-protocol-catalog",
        "protocol catalog index",
        Optional.of("gridgrind --print-protocol-catalog"),
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
        "print-protocol-catalog",
        "protocol catalog search report",
        Optional.of(
            "gridgrind --print-protocol-catalog --search \"" + command.searchQuery() + "\""),
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
    List<String> matches = GridGrindProtocolCatalog.matchingLookupIds(command.lookupId());
    if (matches.size() > 1) {
      String message =
          "Ambiguous lookup id: "
              + command.lookupId()
              + ". Use one of: "
              + String.join(", ", matches);
      return CliCatalogPayloadSupport.writeCliDiagnostic(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CliDiagnostics.invalidArguments(
              2, "print-protocol-catalog", Optional.of("--lookup"), message, matches),
          prettyJson);
    }
    var lookupValue = GridGrindProtocolCatalog.lookupValueFor(command.lookupId());
    if (lookupValue.isEmpty()) {
      String message = CliCatalogCommandSupport.unknownOperationMessage(command.lookupId());
      return CliCatalogPayloadSupport.writeCliDiagnostic(
          responseWriter,
          command.responsePath(),
          stdout,
          stderr,
          CliDiagnostics.invalidArguments(
              2,
              "print-protocol-catalog",
              Optional.of("--lookup"),
              message,
              CliSuggestionSupport.protocolCatalogSearchCommandForLookupId(command.lookupId())
                  .map(List::of)
                  .orElse(List.of())),
          prettyJson);
    }
    return CliCatalogPayloadSupport.writeRenderedPayload(
        responseWriter,
        "print-protocol-catalog",
        "protocol catalog lookup result",
        Optional.of("gridgrind --print-protocol-catalog --lookup " + command.lookupId()),
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
