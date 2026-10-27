package dev.erst.gridgrind.cli.command;

import java.util.Objects;

/** Signals an unrecognized or malformed CLI argument. */
final class CliArgumentsException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  private final String argument;

  CliArgumentsException(String argument, String message) {
    super(message);
    this.argument = Objects.requireNonNull(argument, "argument must not be null");
  }

  CliArgumentsException(String argument, String message, Throwable cause) {
    super(message, cause);
    this.argument = Objects.requireNonNull(argument, "argument must not be null");
  }

  String argument() {
    return argument;
  }
}
