package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.Preservation;
import java.util.List;

/** Host-acceptance outcomes established after plan execution and before publication. */
record HostAcceptanceVerification(
    List<AssertionResult> assertions, List<Preservation> preservation) {
  HostAcceptanceVerification {
    assertions = List.copyOf(assertions);
    preservation = List.copyOf(preservation);
  }

  static HostAcceptanceVerification minimumOnly() {
    return new HostAcceptanceVerification(List.of(), List.of());
  }
}
