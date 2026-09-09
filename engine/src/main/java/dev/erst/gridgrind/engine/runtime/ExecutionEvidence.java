package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookOperationContracts;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Projects only facts established by the execution workflow into the result evidence record. */
final class ExecutionEvidence {
  private ExecutionEvidence() {}

  static WorkbookExecutionEvidence forExecution(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      CalculationReport calculation,
      List<AssertionResult> planAssertions,
      List<AssertionResult> hostAssertions,
      List<WorkbookExecutionEvidence.Preservation> preservation,
      boolean stagedArtifactVerified) {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(bindings, "bindings must not be null");
    return new WorkbookExecutionEvidence(
        admission(request, bindings),
        stagedArtifactVerified
            ? new WorkbookExecutionEvidence.Structural.Established()
            : new WorkbookExecutionEvidence.Structural.NotEstablished(),
        computational(calculation),
        taskSpecific(planAssertions, hostAssertions),
        new WorkbookExecutionEvidence.Presentational.NotAssessed(),
        preservationOrNotRequested(preservation));
  }

  static WorkbookExecutionEvidence forExecution(
      WorkbookPlan request,
      CalculationReport calculation,
      List<AssertionResult> planAssertions,
      List<AssertionResult> hostAssertions,
      List<WorkbookExecutionEvidence.Preservation> preservation,
      boolean stagedArtifactVerified) {
    Objects.requireNonNull(request, "request must not be null");
    return forExecution(
        request,
        WorkbookExecutionEvidence.Admission.empty(),
        calculation,
        planAssertions,
        hostAssertions,
        preservation,
        stagedArtifactVerified);
  }

  private static WorkbookExecutionEvidence forExecution(
      WorkbookPlan request,
      WorkbookExecutionEvidence.Admission admission,
      CalculationReport calculation,
      List<AssertionResult> planAssertions,
      List<AssertionResult> hostAssertions,
      List<WorkbookExecutionEvidence.Preservation> preservation,
      boolean stagedArtifactVerified) {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(admission, "admission must not be null");
    Objects.requireNonNull(calculation, "calculation must not be null");
    return new WorkbookExecutionEvidence(
        admission,
        stagedArtifactVerified
            ? new WorkbookExecutionEvidence.Structural.Established()
            : new WorkbookExecutionEvidence.Structural.NotEstablished(),
        computational(calculation),
        taskSpecific(planAssertions, hostAssertions),
        new WorkbookExecutionEvidence.Presentational.NotAssessed(),
        preservationOrNotRequested(preservation));
  }

  private static WorkbookExecutionEvidence.Admission admission(
      WorkbookPlan request, ExecutionInputBindings bindings) {
    List<WorkbookExecutionEvidence.Operation> operations =
        request.steps().stream().map(ExecutionEvidence::operation).toList();
    List<WorkbookExecutionEvidence.MaterializedInput> inputs =
        bindings.hasRequestPathAccess()
            ? bindings.requestPathAccess().materializedReadIdentities().stream()
                .map(
                    input ->
                        new WorkbookExecutionEvidence.MaterializedInput(
                            input.role(), input.path(), input.byteSize(), input.sha256()))
                .toList()
            : List.of();
    return new WorkbookExecutionEvidence.Admission(operations, inputs);
  }

  private static WorkbookExecutionEvidence.Operation operation(WorkbookStep step) {
    Object operation =
        switch (step) {
          case MutationStep mutation -> mutation.action();
          case AssertionStep assertion -> assertion.assertion();
          case InspectionStep inspection -> inspection.query();
        };
    var semantics = WorkbookOperationContracts.semanticsFor(operation);
    return new WorkbookExecutionEvidence.Operation(
        switch (step) {
          case MutationStep mutation -> mutation.action().actionType();
          case AssertionStep assertion -> assertion.assertion().assertionType();
          case InspectionStep inspection -> inspection.query().queryType();
        },
        semantics.effects(),
        semantics.footprint());
  }

  private static WorkbookExecutionEvidence.Computational computational(
      CalculationReport calculation) {
    return switch (calculation.execution().status()) {
      case NOT_REQUESTED -> new WorkbookExecutionEvidence.Computational.NotRequested();
      case SUCCEEDED -> new WorkbookExecutionEvidence.Computational.Established();
      case SKIPPED, PARTIAL, FAILED -> new WorkbookExecutionEvidence.Computational.Unestablished();
    };
  }

  private static WorkbookExecutionEvidence.TaskSpecific taskSpecific(
      List<AssertionResult> planAssertions, List<AssertionResult> hostAssertions) {
    List<WorkbookExecutionEvidence.TaskSpecific.Assertion> assertions = new ArrayList<>();
    for (AssertionResult assertion :
        List.copyOf(Objects.requireNonNull(planAssertions, "planAssertions must not be null"))) {
      assertions.add(
          new WorkbookExecutionEvidence.TaskSpecific.Assertion(
              WorkbookExecutionEvidence.TaskSpecific.Origin.PLAN_AUTHORED, assertion));
    }
    for (AssertionResult assertion :
        List.copyOf(Objects.requireNonNull(hostAssertions, "hostAssertions must not be null"))) {
      assertions.add(
          new WorkbookExecutionEvidence.TaskSpecific.Assertion(
              WorkbookExecutionEvidence.TaskSpecific.Origin.HOST_REQUIRED, assertion));
    }
    if (assertions.isEmpty()) {
      return new WorkbookExecutionEvidence.TaskSpecific.NotRequested();
    }
    return assertions.stream()
            .anyMatch(assertion -> assertion.result() instanceof AssertionResult.Failed)
        ? new WorkbookExecutionEvidence.TaskSpecific.Failed(assertions)
        : new WorkbookExecutionEvidence.TaskSpecific.Established(assertions);
  }

  private static List<WorkbookExecutionEvidence.Preservation> preservationOrNotRequested(
      List<WorkbookExecutionEvidence.Preservation> preservation) {
    List<WorkbookExecutionEvidence.Preservation> copied =
        List.copyOf(Objects.requireNonNull(preservation, "preservation must not be null"));
    return copied.isEmpty()
        ? List.of(new WorkbookExecutionEvidence.Preservation.NotRequested())
        : copied;
  }
}
