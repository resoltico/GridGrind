package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Requests that a minimal fail-closed host-grant JSON document be emitted as primary output. */
public record CliGrantTemplateCommand(Optional<Path> responsePath) implements CliCommand {
  public CliGrantTemplateCommand {
    Objects.requireNonNull(responsePath, "responsePath must not be null");
  }
}
