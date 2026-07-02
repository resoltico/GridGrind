package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Parses the protocol-catalog immediate command family. */
final class CliProtocolCatalogCommandParser {
  private CliProtocolCatalogCommandParser() {}

  static CliImmediateCommandParser.Result parse(
      String[] args, int index, Optional<Path> responsePath) {
    Objects.requireNonNull(args, "args must not be null");
    Objects.requireNonNull(responsePath, "responsePath must not be null");
    ProtocolCatalogSelection selection = new ProtocolCatalogSelection();
    int nextIndex = index + 1;
    while (nextIndex < args.length) {
      int consumedIndex = selection.tryConsume(args, nextIndex);
      if (consumedIndex == nextIndex) {
        return selection.command(responsePath, nextIndex);
      }
      nextIndex = consumedIndex;
    }
    return selection.command(responsePath, nextIndex);
  }

  /** Tracks the mutually exclusive protocol-catalog subcommand flags while parsing argv. */
  private static final class ProtocolCatalogSelection {
    private Optional<String> lookupFilter = Optional.empty();
    private Optional<String> searchQuery = Optional.empty();

    private int tryConsume(String[] args, int index) {
      String argument = args[index];
      if ("--lookup".equals(argument)) {
        return consumeLookup(args, index);
      }
      if ("--search".equals(argument)) {
        return consumeSearch(args, index);
      }
      return index;
    }

    private int consumeLookup(String[] args, int index) {
      if (lookupFilter.isPresent()) {
        throw new CliArgumentsException("--lookup", "Duplicate argument: --lookup");
      }
      int valueIndex = CliPathArguments.nextValueIndex(args, index, "--lookup");
      lookupFilter =
          Optional.of(
              CliPathArguments.requireNonBlankValue(
                  "--lookup", args[valueIndex], "protocol catalog lookup id"));
      return valueIndex + 1;
    }

    private int consumeSearch(String[] args, int index) {
      if (searchQuery.isPresent()) {
        throw new CliArgumentsException("--search", "Duplicate argument: --search");
      }
      int valueIndex = CliPathArguments.nextValueIndex(args, index, "--search");
      searchQuery =
          Optional.of(
              CliPathArguments.requireNonBlankValue("--search", args[valueIndex], "search query"));
      return valueIndex + 1;
    }

    private CliImmediateCommandParser.Result command(Optional<Path> responsePath, int nextIndex) {
      if (lookupFilter.isPresent() && searchQuery.isPresent()) {
        throw new CliArgumentsException(
            "--search", "--print-protocol-catalog does not allow both --lookup and --search");
      }
      return new CliImmediateCommandParser.Result(
          command(responsePath), nextIndex, "--print-protocol-catalog");
    }

    private CliCommand.PrintProtocolCatalog command(Optional<Path> responsePath) {
      if (searchQuery.isPresent()) {
        return new CliCommand.PrintProtocolCatalogSearch(searchQuery.orElseThrow(), responsePath);
      }
      if (lookupFilter.isPresent()) {
        return new CliCommand.PrintProtocolCatalogLookup(lookupFilter.orElseThrow(), responsePath);
      }
      return new CliCommand.PrintProtocolCatalogIndex(responsePath);
    }
  }
}
