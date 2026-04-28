package dev.erst.gridgrind.cli;

import java.util.Objects;

/** Signals an unrecognized or malformed CLI argument. */
final class CliArgumentsException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  private final String argument;

  CliArgumentsException(String argument, String message) {
    super(message);
    this.argument = Objects.requireNonNull(argument, "argument must not be null");
  }

  String argument() {
    return argument;
  }
}
