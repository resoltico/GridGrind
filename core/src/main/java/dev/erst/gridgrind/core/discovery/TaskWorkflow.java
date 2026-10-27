package dev.erst.gridgrind.core.discovery;

import java.util.List;

/** Internal grouped workflow structure for one task descriptor. */
public record TaskWorkflow(List<TaskPhase> phases, List<String> commonPitfalls) {
  public TaskWorkflow {
    phases = DiscoveryValidation.copyTaskPhases(phases, "phases");
    commonPitfalls = DiscoveryValidation.copyStrings(commonPitfalls, "commonPitfalls");
    if (phases.isEmpty()) {
      throw new IllegalArgumentException("phases must not be empty");
    }
  }
}
