package dev.erst.gridgrind.fuzz.protocol;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.engine.runtime.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.fuzz.support.FuzzFilesystem;
import dev.erst.gridgrind.fuzz.support.FuzzWorkbookIo;
import dev.erst.gridgrind.fuzz.support.GeneratedProtocolWorkflow;
import dev.erst.gridgrind.fuzz.support.GridGrindFuzzData;
import dev.erst.gridgrind.fuzz.support.OperationSequenceModel;
import dev.erst.gridgrind.fuzz.support.WorkbookInvariantChecks;

/** Fuzzes ordered protocol workflows against the production service entrypoint. */
public final class ProtocolWorkflowHarness {
  private ProtocolWorkflowHarness() {}

  public static void check(GridGrindFuzzData fuzzData) {
    GeneratedProtocolWorkflow workflow;
    try {
      workflow = OperationSequenceModel.nextProtocolWorkflow(fuzzData);
    } catch (IllegalArgumentException expected) {
      return;
    } catch (java.io.IOException unexpected) {
      throw new IllegalStateException(unexpected);
    }
    WorkbookPlan request = workflow.request();
    try {
      WorkbookResult response =
          new DefaultGridGrindRequestExecutor()
              .execute(
                  request, FuzzWorkbookIo.executionBindings(workflow.executionRoot(), request));
      if (response instanceof WorkbookResult.Failure failure
          && failure.problem().code()
              == dev.erst.gridgrind.core.problem.GridGrindProblemCode.AUTHORITY_DENIED) {
        throw new IllegalStateException(
            "Generated workflow was stopped by incomplete fixture authority");
      }
      WorkbookInvariantChecks.requireResponseShape(response);
      WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response);
    } catch (RuntimeException | Error failure) {
      workflow.cleanupRoots().forEach(root -> FuzzFilesystem.cleanupAfterFailure(root, failure));
      throw failure;
    }
    workflow.cleanup();
  }
}
