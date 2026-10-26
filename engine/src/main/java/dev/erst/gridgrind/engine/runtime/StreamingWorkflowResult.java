package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.result.WorkbookResultPersistence;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.util.List;

/** Assembles the single successful response for an admitted streaming-write workflow. */
final class StreamingWorkflowResult {
  private StreamingWorkflowResult() {}

  static WorkbookResult success(
      ProtocolVersion protocolVersion,
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
