package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Machine-readable CLI failure report for argument, lookup, and command-usage problems. */
public record CliFailureReport(
    GridGrindProtocolVersion protocolVersion,
    int exitCode,
    String command,
    GridGrindProblemCode code,
    String message,
    CliFailureLocation location,
    Optional<String> argument,
    List<String> suggestions,
    Optional<String> resolution) {
  public CliFailureReport {
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    if (exitCode <= 0) {
      throw new IllegalArgumentException("exitCode must be positive");
    }
    command = CliDiscoveryValidation.requireNonBlank(command, "command");
    Objects.requireNonNull(code, "code must not be null");
    message = CliDiscoveryValidation.requireNonBlank(message, "message");
    Objects.requireNonNull(location, "location must not be null");
    argument = CliDiscoveryValidation.normalizeOptionalString(argument, "argument");
    suggestions = CliDiscoveryValidation.copyStringsAllowEmpty(suggestions, "suggestions");
    resolution = CliDiscoveryValidation.normalizeOptionalString(resolution, "resolution");
  }
}
