package dev.erst.gridgrind.contract.step;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.catalog.GridGrindExecutionModeMetadata;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;

/** Validates the cross-step sequencing facts unique to STREAMING_WRITE. */
final class WorkbookStaticStreamingValidation {
  private static final GridGrindExecutionModeMetadata.StreamingWriteMode STREAMING_WRITE =
      GridGrindExecutionModeMetadata.streamingWrite();

  private WorkbookStaticStreamingValidation() {}

  static void add(
      List<WorkbookStaticViolation> violations,
      WorkbookStaticRequest request,
      ExecutionPolicyInput execution) {
    if (!execution.calculation().allowsStreamingWrite()) {
      violations.add(
          new WorkbookStaticViolation(
              "execution.mode", STREAMING_WRITE.calculationFailureMessage()));
    }
    request
        .source()
        .filter(source -> !(source instanceof WorkbookPlan.WorkbookSource.New))
        .ifPresent(
            source ->
                violations.add(
                    new WorkbookStaticViolation(
                        "execution.mode", STREAMING_WRITE.invalidSourceMessage())));
    boolean seenEnsureSheet = false;
    boolean everyStepBound = request.steps().stream().allMatch(step -> step.value().isPresent());
    for (int index = 0; index < request.steps().size(); index++) {
      WorkbookStaticStep step = request.steps().get(index);
      if (step.value().isEmpty()) {
        continue;
      }
      WorkbookStep value = step.value().orElseThrow();
      addStepViolations(
          violations, value, seenEnsureSheet, priorStepsBound(request.steps(), index));
      seenEnsureSheet |= isEnsureSheet(value);
    }
    if (everyStepBound && !seenEnsureSheet) {
      violations.add(
          new WorkbookStaticViolation(
              "execution.mode", STREAMING_WRITE.missingEnsureSheetMutationMessage()));
    }
  }

  private static void addStepViolations(
      List<WorkbookStaticViolation> violations,
      WorkbookStep step,
      boolean seenEnsureSheet,
      boolean everyPriorStepBound) {
    if (step instanceof MutationStep mutation) {
      WorkbookOperationContracts.executionModeViolation(
              mutation.action(), ExecutionModeInput.streamingWrite())
          .ifPresent(
              message -> violations.add(new WorkbookStaticViolation("execution.mode", message)));
      if (mutation.action() instanceof CellMutationAction.AppendRow
          && !seenEnsureSheet
          && everyPriorStepBound) {
        violations.add(
            new WorkbookStaticViolation(
                "execution.mode", STREAMING_WRITE.missingEnsureSheetBeforeAppendMessage()));
      }
      return;
    }
    if (!seenEnsureSheet && everyPriorStepBound) {
      String message =
          step instanceof AssertionStep
              ? STREAMING_WRITE.missingEnsureSheetBeforeAssertionMessage()
              : STREAMING_WRITE.missingEnsureSheetBeforeInspectionMessage();
      violations.add(new WorkbookStaticViolation("execution.mode", message));
    }
  }

  private static boolean priorStepsBound(List<WorkbookStaticStep> steps, int upperExclusive) {
    for (int index = 0; index < upperExclusive; index++) {
      if (steps.get(index).value().isEmpty()) {
        return false;
      }
    }
    return true;
  }

  private static boolean isEnsureSheet(WorkbookStep step) {
    return step instanceof MutationStep mutation
        && mutation.action() instanceof WorkbookMutationAction.EnsureSheet;
  }
}
