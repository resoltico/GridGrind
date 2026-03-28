package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.Objects;

/** Parses raw command-line arguments into a typed GridGrind CLI command. */
final class CliArguments {
  private CliArguments() {}

  /** Parses the raw CLI args into the corresponding command model. */
  static CliCommand parse(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    Path requestPath = null;
    Path responsePath = null;

    int index = 0;
    while (index < args.length) {
      String argument = args[index];
      switch (argument) {
        case "--help", "-h" -> {
          return new CliCommand.Help();
        }
        case "--version" -> {
          return new CliCommand.Version();
        }
        case "--print-request-template" -> {
          return new CliCommand.PrintRequestTemplate();
        }
        case "--print-protocol-catalog" -> {
          return new CliCommand.PrintProtocolCatalog();
        }
        case "--request" -> {
          if (requestPath != null) {
            throw new CliArgumentsException("--request", "Duplicate argument: --request");
          }
          int valueIndex = nextValueIndex(args, index, "--request");
          requestPath = Path.of(args[valueIndex]);
          index = valueIndex + 1;
        }
        case "--response" -> {
          if (responsePath != null) {
            throw new CliArgumentsException("--response", "Duplicate argument: --response");
          }
          int valueIndex = nextValueIndex(args, index, "--response");
          responsePath = Path.of(args[valueIndex]);
          index = valueIndex + 1;
        }
        default -> throw new CliArgumentsException(argument, "Unknown argument: " + argument);
      }
    }

    return new CliCommand.Execute(requestPath, responsePath);
  }

  private static int nextValueIndex(String[] args, int flagIndex, String flagName) {
    int valueIndex = flagIndex + 1;
    if (valueIndex >= args.length) {
      throw new CliArgumentsException(flagName, "Missing value for " + flagName);
    }
    return valueIndex;
  }
}
