package dev.erst.gridgrind.cli;

import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/** Shared primary-command detection for raw CLI argv handling. */
final class CliPrimaryCommandSupport {
  private static final Map<String, String> PRIMARY_COMMAND_NAMES =
      Map.ofEntries(
          Map.entry("--help", "help"),
          Map.entry("-h", "help"),
          Map.entry("--help-protocol", "help-protocol"),
          Map.entry("--help-guidance", "help-guidance"),
          Map.entry("--version", "version"),
          Map.entry("--license", "license"),
          Map.entry("--print-request-template", "print-request-template"),
          Map.entry("--print-example", "print-example"),
          Map.entry("--print-example-catalog", "print-example-catalog"),
          Map.entry("--print-task-catalog", "print-task-catalog"),
          Map.entry("--print-task-plan", "print-task-plan"),
          Map.entry("--print-task-keyword-match", "print-task-keyword-match"),
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
    return "execute";
  }

  static boolean isPrimaryCommandToken(String argument) {
    Objects.requireNonNull(argument, "argument must not be null");
    return commandName(argument).isPresent();
  }

  private static Optional<String> commandName(String argument) {
    return Optional.ofNullable(PRIMARY_COMMAND_NAMES.get(argument));
  }
}
