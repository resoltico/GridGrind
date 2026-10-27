package dev.erst.gridgrind.cli.command;

import java.util.Optional;

/** Supplies CLI remediation from the command/argument context, without parsing error prose. */
final class ArgumentGuidance {
  private ArgumentGuidance() {}

  static String resolution(String command, Optional<String> argument) {
    return switch (argument.orElse("")) {
      case "--request" ->
          "Provide a readable request JSON file, or use --request - with an explicit --execution-root.";
      case "--execution-root" ->
          "Provide the explicit directory used to resolve stdin request paths.";
      case "--grant" ->
          "Create template grant, then authorize the intended input paths, targets and publication.";
      case "--response" -> "Choose a writable response path distinct from the request file.";
      default ->
          command.startsWith("recipe")
              ? "Use gridgrind recipe list for IDs, or gridgrind recipe match for a goal."
              : command.startsWith("catalog")
                  ? "Use gridgrind catalog search, then catalog show with an exact qualified ID."
                  : "Use gridgrind help to inspect commands and supported options.";
    };
  }
}
