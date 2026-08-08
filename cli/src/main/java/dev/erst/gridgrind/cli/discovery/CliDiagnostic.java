package dev.erst.gridgrind.cli.discovery;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Machine-readable CLI diagnostic with canonical problem cores plus transport metadata. */
public record CliDiagnostic(
    GridGrindProtocolVersion protocolVersion,
    int exitCode,
    String command,
    List<String> suggestions,
    List<GridGrindProblemDetail.Problem> problems,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CliTransport> transport) {
  public CliDiagnostic {
    protocolVersion = CliDiscoveryValidation.requireProtocolVersion(protocolVersion);
    if (exitCode <= 0) {
      throw new IllegalArgumentException("exitCode must be positive");
    }
    command = CliDiscoveryValidation.requireNonBlank(command, "command");
    suggestions = CliDiscoveryValidation.copyStringsAllowEmpty(suggestions, "suggestions");
    problems = List.copyOf(Objects.requireNonNull(problems, "problems must not be null"));
    if (problems.isEmpty()) {
      throw new IllegalArgumentException("problems must not be empty");
    }
    problems.forEach(problem -> Objects.requireNonNull(problem, "problems must not contain null"));
    transport = CliDiscoveryValidation.copyOptionalTransport(transport, "transport");
  }

  /**
   * Returns the leading diagnostic problem for callers that intentionally present one summary.
   *
   * <p>Machine consumers must use {@link #problems()} so independent intake findings are never
   * discarded.
   */
  public GridGrindProblemDetail.Problem primaryProblem() {
    return problems.getFirst();
  }

  /**
   * Returns the leading diagnostic problem for a human-facing summary.
   *
   * <p>This derived accessor is not part of the wire contract. Machine consumers must use {@link
   * #problems()}.
   */
  public GridGrindProblemDetail.Problem problem() {
    return primaryProblem();
  }
}
