package dev.erst.gridgrind.cli.discovery;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Machine-readable CLI diagnostic with one canonical Problem core plus transport metadata. */
public record CliDiagnostic(
    GridGrindProtocolVersion protocolVersion,
    int exitCode,
    String command,
    List<String> suggestions,
    GridGrindProblemDetail.Problem problem,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CliTransport> transport) {
  public CliDiagnostic {
    protocolVersion = CliDiscoveryValidation.requireProtocolVersion(protocolVersion);
    if (exitCode <= 0) {
      throw new IllegalArgumentException("exitCode must be positive");
    }
    command = CliDiscoveryValidation.requireNonBlank(command, "command");
    suggestions = CliDiscoveryValidation.copyStringsAllowEmpty(suggestions, "suggestions");
    Objects.requireNonNull(problem, "problem must not be null");
    transport = CliDiscoveryValidation.copyOptionalTransport(transport, "transport");
  }
}
