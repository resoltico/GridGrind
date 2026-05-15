package dev.erst.gridgrind.cli.discovery;

import java.util.List;

/** Internal grouped workflow structure for one task descriptor. */
public record TaskWorkflow(List<TaskPhase> phases, List<String> commonPitfalls) {
  public TaskWorkflow {
    phases = CliDiscoveryValidation.copyTaskPhases(phases, "phases");
    commonPitfalls = CliDiscoveryValidation.copyStrings(commonPitfalls, "commonPitfalls");
    if (phases.isEmpty()) {
      throw new IllegalArgumentException("phases must not be empty");
    }
  }
}
