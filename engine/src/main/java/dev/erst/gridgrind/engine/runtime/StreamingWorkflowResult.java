package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import java.util.List;

/** Assembles the single successful response for an admitted streaming-write workflow. */
final class StreamingWorkflowResult {
  private StreamingWorkflowResult() {}

  static WorkbookResult success(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      StreamingWorkflowContext context,
      CalculationReport calculation,
      WorkbookResultPersistence.PersistenceOutcome persistence) {
    return new WorkbookResult.Success(
        protocolVersion,
        request.planId(),
        context.journal().buildSuccess(request.steps().size()),
        calculation,
        persistence,
        ExecutionEvidence.forExecution(
            request,
            bindings,
            calculation,
            context.planAssertions(),
            context.hostAssertions(),
            context.preservation(),
            !(persistence instanceof WorkbookResultPersistence.PersistenceOutcome.NotSaved)),
        context.warnings(),
        List.copyOf(context.planAssertions()),
        List.copyOf(context.inspections()));
  }
}
