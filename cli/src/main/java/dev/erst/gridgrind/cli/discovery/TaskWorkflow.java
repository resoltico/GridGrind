package dev.erst.gridgrind.cli.discovery;

import java.util.List;

/** Internal grouped workflow structure for one task descriptor. */
record TaskWorkflow(List<TaskPhase> phases, List<String> commonPitfalls) {
  TaskWorkflow {
    phases = CliDiscoveryValidation.copyTaskPhases(phases, "phases");
    commonPitfalls = CliDiscoveryValidation.copyStrings(commonPitfalls, "commonPitfalls");
  }
}
