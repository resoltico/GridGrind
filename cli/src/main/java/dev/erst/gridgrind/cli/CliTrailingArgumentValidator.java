package dev.erst.gridgrind.cli;

/** Validates that an immediate CLI command consumed the full argument vector. */
final class CliTrailingArgumentValidator {
  private CliTrailingArgumentValidator() {}

  static void requireNoTrailingArguments(String[] args, int nextIndex, String commandToken) {
    if (nextIndex == args.length) {
      return;
    }
    String trailingArgument = args[nextIndex];
    if (CliPrimaryCommandSupport.isPrimaryCommandToken(trailingArgument)) {
      throw CliPrimaryCommandSupport.multiplePrimaryCommands(commandToken, trailingArgument);
    }
    if ("--request".equals(trailingArgument)
        || "--execution-root".equals(trailingArgument)
        || "--temp-root".equals(trailingArgument)) {
      throw CliPrimaryCommandSupport.commandDoesNotAllowFlag(commandToken, trailingArgument);
    }
    throw CliExecutionCommandParser.unknownArgumentException(trailingArgument);
  }
}
