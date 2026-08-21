package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CommandError;
import java.util.Objects;
import java.util.Optional;

/** Projects argument-parser failures into the canonical rejected-command result. */
final class CliArgumentFailureSupport {
  private CliArgumentFailureSupport() {}

  static CommandError reportFor(String[] args, CliArgumentsException exception) {
    Objects.requireNonNull(args, "args must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    return CommandErrors.invalidArguments(
        CliPrimaryCommandSupport.primaryCommandName(args),
        Optional.of(exception.argument()),
        Objects.requireNonNullElse(exception.getMessage(), "Invalid command-line arguments"));
  }

  static CommandError reportFor(String[] args, IllegalArgumentException exception) {
    Objects.requireNonNull(args, "args must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    return CommandErrors.invalidArguments(
        CliPrimaryCommandSupport.primaryCommandName(args),
        Optional.empty(),
        Objects.requireNonNullElse(exception.getMessage(), "Invalid command-line arguments"));
  }
}
