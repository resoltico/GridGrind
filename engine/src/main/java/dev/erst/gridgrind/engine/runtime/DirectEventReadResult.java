package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import java.util.List;

/** Assembles a successful admitted direct-event-read result without a publication claim. */
final class DirectEventReadResult {
  private DirectEventReadResult() {}

  static WorkbookResult.Success success(
      DirectEventReadContext context, ExecutionInputBindings bindings) {
    return new WorkbookResult.Success(
        context.protocolVersion(),
        context.request().planId(),
        context.journal().buildSuccess(context.request().steps().size(), false),
        context.calculation(),
        new WorkbookResultPersistence.PersistenceOutcome.NotSaved(),
        ExecutionEvidence.forExecution(
            context.request(),
            bindings,
            context.calculation(),
            List.of(),
            context.hostAssertions(),
            context.preservation(),
            false),
        context.warnings(),
        List.of(),
        List.copyOf(context.inspections()));
  }
}
