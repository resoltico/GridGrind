package dev.erst.gridgrind.cli;

import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/** Shared primary-command detection for raw CLI argv handling. */
final class CliPrimaryCommandSupport {
  private static final Map<String, String> PRIMARY_COMMAND_NAMES =
      Map.ofEntries(
          Map.entry("--help", "help"),
          Map.entry("help", "help"),
          Map.entry("-h", "help"),
          Map.entry("--help-protocol", "help-protocol"),
          Map.entry("help-protocol", "help-protocol"),
          Map.entry("--help-guidance", "help-guidance"),
          Map.entry("help-guidance", "help-guidance"),
          Map.entry("--version", "version"),
          Map.entry("version", "version"),
          Map.entry("--license", "license"),
          Map.entry("license", "license"),
          Map.entry("--print-request-template", "print-request-template"),
          Map.entry("--print-recipe", "print-recipe"),
          Map.entry("--materialize-recipe", "materialize-recipe"),
          Map.entry("--print-recipe-catalog", "print-recipe-catalog"),
          Map.entry("--print-recipe-keyword-match", "print-recipe-keyword-match"),
          Map.entry("--print-protocol-catalog", "print-protocol-catalog"),
          Map.entry("--doctor-request", "doctor-request"));

  private CliPrimaryCommandSupport() {}

  static String primaryCommandName(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    for (String argument : args) {
      Optional<String> commandName = commandName(argument);
      if (commandName.isPresent()) {
        return commandName.orElseThrow();
      }
    }
    return "cli";
  }

  static boolean isPrimaryCommandToken(String argument) {
    Objects.requireNonNull(argument, "argument must not be null");
    return commandName(argument).isPresent();
  }

  static CliArgumentsException multiplePrimaryCommands(
      String firstCommandToken, String secondToken) {
    Objects.requireNonNull(firstCommandToken, "firstCommandToken must not be null");
    Objects.requireNonNull(secondToken, "secondToken must not be null");
    return new CliArgumentsException(
        secondToken,
        "Only one primary command may be used per invocation; "
            + firstCommandToken
            + " cannot be combined with "
            + secondToken);
  }

  static CliArgumentsException commandDoesNotAllowFlag(String commandToken, String flagName) {
    Objects.requireNonNull(commandToken, "commandToken must not be null");
    Objects.requireNonNull(flagName, "flagName must not be null");
    return new CliArgumentsException(flagName, commandToken + " does not allow " + flagName);
  }

  static CliArgumentsException commandMustBePrimary(String commandToken) {
    Objects.requireNonNull(commandToken, "commandToken must not be null");
    return new CliArgumentsException(
        commandToken,
        commandToken + " must be the primary command and cannot follow execution arguments");
  }

  private static Optional<String> commandName(String argument) {
    return Optional.ofNullable(PRIMARY_COMMAND_NAMES.get(argument));
  }
}
