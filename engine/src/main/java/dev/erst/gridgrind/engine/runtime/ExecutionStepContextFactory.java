package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.StepReference;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.step.WorkbookStep;

/** Creates consistent step references and diagnostic contexts for execution failures. */
final class ExecutionStepContextFactory {
  private ExecutionStepContextFactory() {}

  static ProblemContext.ExecuteStep contextFor(
      WorkbookPlan request, int stepIndex, WorkbookStep step, Exception exception) {
    return new ProblemContext.ExecuteStep(
        ExecutionRequestPaths.requestShape(request),
        referenceFor(stepIndex, step),
        ExecutionDiagnosticFields.locationFor(step, exception));
  }

  static StepReference referenceFor(int stepIndex, WorkbookStep step) {
    return new StepReference(
        stepIndex, step.stepId(), step.stepKind(), ExecutionStepKinds.stepType(step));
  }
}
