package dev.erst.gridgrind.cli.discovery;

import com.fasterxml.jackson.annotation.JsonIgnore;
import dev.erst.gridgrind.core.problem.DiagnosticOrder;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.util.List;
import java.util.Objects;

/** Canonical plural result for a command rejected before workbook execution begins. */
public record CommandError(
    ProtocolVersion protocolVersion,
    String command,
    String status,
    List<GridGrindProblemDetail.Problem> problems) {
  /** Constructs a rejection with the fixed command status. */
  public CommandError(
      ProtocolVersion protocolVersion,
      String command,
      List<GridGrindProblemDetail.Problem> problems) {
    this(protocolVersion, command, "REJECTED", problems);
  }

  public CommandError {
    if (!"REJECTED".equals(status)) {
      throw new IllegalArgumentException("CommandError status must be REJECTED");
    }
    protocolVersion = CliDiscoveryValidation.requireProtocolVersion(protocolVersion);
    command = CliDiscoveryValidation.requireNonBlank(command, "command");
    problems =
        DiagnosticOrder.problems(
            List.copyOf(Objects.requireNonNull(problems, "problems must not be null")));
    if (problems.isEmpty()) {
      throw new IllegalArgumentException("problems must not be empty");
    }
  }

  /** Returns the leading problem for narrow human summaries without changing the wire shape. */
  @JsonIgnore
  public GridGrindProblemDetail.Problem primaryProblem() {
    return problems.getFirst();
  }
}
