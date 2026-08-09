package dev.erst.gridgrind.cli.discovery;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.contract.dto.DiagnosticOrder;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.List;
import java.util.Objects;

/** Canonical plural result for a command rejected before workbook execution begins. */
public record CommandError(
    GridGrindProtocolVersion protocolVersion,
    String command,
    List<GridGrindProblemDetail.Problem> problems) {
  public CommandError {
    protocolVersion = CliDiscoveryValidation.requireProtocolVersion(protocolVersion);
    command = CliDiscoveryValidation.requireNonBlank(command, "command");
    problems =
        DiagnosticOrder.problems(
            List.copyOf(Objects.requireNonNull(problems, "problems must not be null")));
    if (problems.isEmpty()) {
      throw new IllegalArgumentException("problems must not be empty");
    }
  }

  /**
   * Recreates one rejected-command result only when its fixed wire status is present and REJECTED.
   */
  @JsonCreator
  public static CommandError fromJson(
      @JsonProperty("protocolVersion") GridGrindProtocolVersion protocolVersion,
      @JsonProperty("command") String command,
      @JsonProperty("status") String status,
      @JsonProperty("problems") List<GridGrindProblemDetail.Problem> problems) {
    if (!"REJECTED".equals(status)) {
      throw new IllegalArgumentException("CommandError status must be REJECTED");
    }
    return new CommandError(protocolVersion, command, problems);
  }

  /** The command never began workbook execution. */
  @JsonProperty("status")
  public String status() {
    return "REJECTED";
  }

  /** Returns the leading problem for narrow human summaries without changing the wire shape. */
  @JsonIgnore
  public GridGrindProblemDetail.Problem primaryProblem() {
    return problems.getFirst();
  }
}
