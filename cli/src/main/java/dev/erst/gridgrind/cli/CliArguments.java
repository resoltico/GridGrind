package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Parses raw command-line arguments into a typed GridGrind CLI command. */
final class CliArguments {
  private CliArguments() {}

  /** Parses the raw CLI args into the corresponding command model. */
  static CliCommand parse(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    CliPathArguments.GlobalResponseExtraction extraction =
        CliPathArguments.extractGlobalResponse(args);
    String[] remainingArgs = extraction.remainingArgsArray();
    Optional<Path> responsePath = extraction.responsePath();
    if (remainingArgs.length > 0) {
      Optional<CliImmediateCommandParser.Result> immediate =
          CliImmediateCommandParser.parse(remainingArgs, 0, remainingArgs[0], responsePath);
      if (immediate.isPresent()) {
        CliImmediateCommandParser.Result result = immediate.orElseThrow();
        requireNoTrailingArguments(remainingArgs, result.nextIndex(), result.commandToken());
        return result.command();
      }
    }
    return CliExecutionCommandParser.parse(remainingArgs, responsePath);
  }

  private static void requireNoTrailingArguments(
      String[] args, int nextIndex, String commandToken) {
    if (nextIndex == args.length) {
      return;
    }
    String trailingArgument = args[nextIndex];
    if (CliPrimaryCommandSupport.isPrimaryCommandToken(trailingArgument)) {
      throw new CliArgumentsException(
          trailingArgument,
          "Only one primary command may be used per invocation; "
              + commandToken
              + " cannot be combined with "
              + trailingArgument);
    }
    if ("--request".equals(trailingArgument)
        || "--execution-root".equals(trailingArgument)
        || "--temp-root".equals(trailingArgument)) {
      throw new CliArgumentsException(
          trailingArgument, commandToken + " does not allow " + trailingArgument);
    }
    throw CliExecutionCommandParser.unknownArgumentException(trailingArgument);
  }
}
