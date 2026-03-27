package dev.erst.gridgrind.cli;

/** Signals an unrecognized or malformed CLI argument. */
final class CliArgumentsException extends IllegalArgumentException {
  private static final long serialVersionUID = 1L;

  private final String argument;

  CliArgumentsException(String argument, String message) {
    super(message);
    this.argument = argument;
  }

  String argument() {
    return argument;
  }
}
