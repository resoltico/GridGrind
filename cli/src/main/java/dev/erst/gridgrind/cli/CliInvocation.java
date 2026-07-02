package dev.erst.gridgrind.cli;

import java.util.Objects;
import java.util.Optional;

/** Parsed CLI invocation including the primary command and global render options. */
record CliInvocation(
    CliCommand command, Optional<CliOutputFormat> outputFormat, boolean prettyJson) {
  CliInvocation {
    Objects.requireNonNull(command, "command must not be null");
    Objects.requireNonNull(outputFormat, "outputFormat must not be null");
  }
}
