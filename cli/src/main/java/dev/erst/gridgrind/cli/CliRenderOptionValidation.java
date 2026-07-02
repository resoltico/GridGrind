package dev.erst.gridgrind.cli;

import java.util.Objects;
import java.util.Optional;

/** Validates authored render options against the owning CLI command family. */
final class CliRenderOptionValidation {
  private static final String STRUCTURED_TEXT_COMMANDS =
      "--help, --help-protocol, --help-guidance, --version, or --license";

  private CliRenderOptionValidation() {}

  static void validate(CliCommand command, Optional<CliOutputFormat> outputFormat) {
    Objects.requireNonNull(command, "command must not be null");
    Objects.requireNonNull(outputFormat, "outputFormat must not be null");
    if (outputFormat.isEmpty() || allowsStructuredTextFormat(command)) {
      return;
    }
    throw new CliArgumentsException(
        "--format",
        "--format is only valid with "
            + STRUCTURED_TEXT_COMMANDS
            + "; JSON-native commands already emit JSON and use --pretty when indentation is"
            + " desired");
  }

  private static boolean allowsStructuredTextFormat(CliCommand command) {
    return command instanceof CliCommand.Help
        || command instanceof CliCommand.Version
        || command instanceof CliCommand.License;
  }
}
