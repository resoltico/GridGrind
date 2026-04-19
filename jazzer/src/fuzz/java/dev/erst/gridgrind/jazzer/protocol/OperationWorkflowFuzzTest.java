package dev.erst.gridgrind.jazzer.protocol;

import com.code_intelligence.jazzer.api.FuzzedDataProvider;
import com.code_intelligence.jazzer.junit.FuzzTest;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.executor.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.jazzer.support.GeneratedProtocolWorkflow;
import dev.erst.gridgrind.jazzer.support.GridGrindFuzzData;
import dev.erst.gridgrind.jazzer.support.HarnessTelemetry;
import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import dev.erst.gridgrind.jazzer.support.OperationSequenceModel;
import dev.erst.gridgrind.jazzer.support.SequenceIntrospection;
import dev.erst.gridgrind.jazzer.support.WorkbookInvariantChecks;

/** Fuzzes ordered protocol workflows against the production service entrypoint. */
class OperationWorkflowFuzzTest {
  private static final HarnessTelemetry TELEMETRY =
      HarnessTelemetry.forHarness(JazzerHarness.protocolWorkflow());

  @FuzzTest
  void executeWorkflow(FuzzedDataProvider data) {
    GridGrindFuzzData fuzzData = GridGrindFuzzData.wrap(data);
    TELEMETRY.beginIteration(fuzzData.remainingBytes());
    GeneratedProtocolWorkflow workflow;
    try {
      workflow = OperationSequenceModel.nextProtocolWorkflow(fuzzData);
    } catch (IllegalArgumentException expected) {
      TELEMETRY.recordGeneratedInvalid();
      return;
    } catch (java.io.IOException unexpected) {
      TELEMETRY.recordUnexpectedFailure(unexpected);
      throw new IllegalStateException(unexpected);
    }
    WorkbookPlan request = workflow.request();
    TELEMETRY.recordSourceKind(SequenceIntrospection.sourceKind(request));
    TELEMETRY.recordPersistenceKind(SequenceIntrospection.persistenceKind(request));
    TELEMETRY.recordSequenceKinds(SequenceIntrospection.mutationKinds(request.mutationSteps()));
    TELEMETRY.recordAssertionKinds(SequenceIntrospection.assertionKinds(request.assertionSteps()));
    TELEMETRY.recordReadKinds(SequenceIntrospection.inspectionKinds(request.inspectionSteps()));
    TELEMETRY.recordStyleKinds(SequenceIntrospection.styleKinds(request.mutationSteps()));
    try {
      try {
        GridGrindResponse response = new DefaultGridGrindRequestExecutor().execute(request);
        WorkbookInvariantChecks.requireResponseShape(response);
        WorkbookInvariantChecks.requireWorkflowOutcomeShape(request, response);
        TELEMETRY.recordResponseKind(SequenceIntrospection.responseKind(response));
        TELEMETRY.recordSuccess();
      } catch (RuntimeException unexpected) {
        TELEMETRY.recordUnexpectedFailure(unexpected);
        throw unexpected;
      }
    } finally {
      workflow.cleanup();
    }
  }
}
