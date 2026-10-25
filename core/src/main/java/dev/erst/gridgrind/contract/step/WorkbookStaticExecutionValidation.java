package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.catalog.GridGrindExecutionModeMetadata;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/** Validates static execution-policy compatibility without interpreting workbook state. */
final class WorkbookStaticExecutionValidation {
  private static final GridGrindExecutionModeMetadata.EventReadMode EVENT_READ =
      GridGrindExecutionModeMetadata.eventRead();

  private WorkbookStaticExecutionValidation() {}

  static List<WorkbookStaticViolation> validate(WorkbookStaticRequest request) {
    if (request.execution().isEmpty()) {
      return List.of();
    }
    ExecutionPolicyInput execution = request.execution().orElseThrow();
    List<WorkbookStaticViolation> violations = new ArrayList<>();
    addCalculationOrderingViolation(violations, request.steps(), execution);
    if (execution.mode() instanceof ExecutionModeInput.EventRead) {
      addEventReadViolations(violations, request.steps(), execution);
    } else if (execution.mode() instanceof ExecutionModeInput.StreamingWrite) {
      WorkbookStaticStreamingValidation.add(violations, request, execution);
    }
    return List.copyOf(violations);
  }

  private static void addCalculationOrderingViolation(
      List<WorkbookStaticViolation> violations,
      List<WorkbookStaticStep> steps,
      ExecutionPolicyInput execution) {
    if (!execution.calculation().requiresMutationPrefix()) {
      return;
    }
    boolean seenObservation = false;
    for (WorkbookStaticStep step : steps) {
      if (step.value().isEmpty()) {
        continue;
      }
      if (step.value().orElseThrow() instanceof MutationStep) {
        if (seenObservation) {
          violations.add(
              new WorkbookStaticViolation(
                  "execution.calculation",
                  "execution.calculation.strategy="
                      + execution.calculation().effectiveStrategy().strategyType()
                      + " requires all MUTATION steps to appear before any ASSERTION or INSPECTION"
                      + " step so calculation can run once at the mutation-to-observation boundary"));
          return;
        }
      } else {
        seenObservation = true;
      }
    }
  }

  private static void addEventReadViolations(
      List<WorkbookStaticViolation> violations,
      List<WorkbookStaticStep> steps,
      ExecutionPolicyInput execution) {
    if (!execution.calculation().allowsEventRead()) {
      violations.add(
          new WorkbookStaticViolation("execution.mode", EVENT_READ.calculationFailureMessage()));
    }
    for (WorkbookStaticStep step : steps) {
      step.value()
          .flatMap(value -> executionModeViolation(value, execution.mode()))
          .ifPresent(
              message -> violations.add(new WorkbookStaticViolation("execution.mode", message)));
    }
  }

  private static Optional<String> executionModeViolation(
      WorkbookStep step, ExecutionModeInput executionMode) {
    return switch (step) {
      case MutationStep mutation ->
          WorkbookOperationContracts.executionModeViolation(mutation.action(), executionMode);
      case AssertionStep assertion ->
          WorkbookOperationContracts.executionModeViolation(assertion.assertion(), executionMode);
      case InspectionStep inspection ->
          WorkbookOperationContracts.executionModeViolation(inspection.query(), executionMode);
    };
  }
}
