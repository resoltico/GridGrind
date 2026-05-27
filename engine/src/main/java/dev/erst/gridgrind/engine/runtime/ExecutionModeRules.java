package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.catalog.GridGrindExecutionModeMetadata;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.util.Optional;
import java.util.Set;

/** Execution-mode and calculation-policy validation rules for request execution. */
final class ExecutionModeRules {
  private static final GridGrindExecutionModeMetadata.EventReadMode EVENT_READ =
      GridGrindExecutionModeMetadata.eventRead();
  private static final GridGrindExecutionModeMetadata.StreamingWriteMode STREAMING_WRITE =
      GridGrindExecutionModeMetadata.streamingWrite();
  private static final Set<Class<? extends MutationAction>> STREAMING_WRITE_MUTATION_ACTION_TYPES =
      Set.copyOf(STREAMING_WRITE.allowedActions());

  private static final Set<Class<? extends InspectionQuery>> EVENT_READ_INSPECTION_QUERY_TYPES =
      Set.copyOf(EVENT_READ.allowedQueries());

  private ExecutionModeRules() {}

  static Optional<String> calculationPolicyFailure(WorkbookPlan request) {
    if (!CalculationPolicyExecutor.requiresMutationPrefix(request.calculationPolicy())) {
      return Optional.empty();
    }
    boolean seenObservationStep = false;
    for (WorkbookStep step : request.steps()) {
      if (step instanceof MutationStep) {
        if (seenObservationStep) {
          return Optional.of(
              "execution.calculation.strategy="
                  + request.calculationPolicy().effectiveStrategy().strategyType()
                  + " requires all MUTATION steps to appear before any ASSERTION or INSPECTION"
                  + " step so calculation can run once at the mutation-to-observation boundary");
        }
      } else {
        seenObservationStep = true;
      }
    }
    return Optional.empty();
  }

  static Optional<String> executionModeFailure(
      WorkbookPlan request, ExecutionModeInput executionMode) {
    return switch (executionMode) {
      case ExecutionModeInput.FullXssf _ -> Optional.empty();
      case ExecutionModeInput.EventRead _ -> eventReadFailure(request);
      case ExecutionModeInput.StreamingWrite _ -> streamingWriteFailure(request);
    };
  }

  static boolean directEventReadEligible(WorkbookPlan request, ExecutionModeInput executionMode) {
    return executionMode instanceof ExecutionModeInput.EventRead
        && CalculationPolicyExecutor.allowsEventRead(request.calculationPolicy())
        && request.steps().stream().allMatch(InspectionStep.class::isInstance)
        && request.persistence() instanceof WorkbookPlan.WorkbookPersistence.None
        && request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile;
  }

  static ExecutionModeInput executionMode(WorkbookPlan request) {
    return request.effectiveExecutionMode();
  }

  private static Optional<String> eventReadFailure(WorkbookPlan request) {
    if (!CalculationPolicyExecutor.allowsEventRead(request.calculationPolicy())) {
      return Optional.of(EVENT_READ.calculationFailureMessage());
    }
    for (WorkbookStep step : request.steps()) {
      if (!(step instanceof InspectionStep inspectionStep)) {
        return Optional.of(EVENT_READ.unsupportedStepMessage(step.stepKind()));
      }
      if (!EVENT_READ_INSPECTION_QUERY_TYPES.contains(inspectionStep.query().getClass())) {
        return Optional.of(EVENT_READ.unsupportedQueryMessage(inspectionStep.query().queryType()));
      }
    }
    return Optional.empty();
  }

  private static Optional<String> streamingWriteFailure(WorkbookPlan request) {
    if (!CalculationPolicyExecutor.allowsStreamingWrite(request.calculationPolicy())) {
      return Optional.of(STREAMING_WRITE.calculationFailureMessage());
    }
    if (!(request.source() instanceof WorkbookPlan.WorkbookSource.New)) {
      return Optional.of(STREAMING_WRITE.invalidSourceMessage());
    }
    boolean seenEnsureSheet = false;
    for (WorkbookStep step : request.steps()) {
      Optional<String> failure = streamingWriteStepFailure(step, seenEnsureSheet);
      if (failure.isPresent()) {
        return failure;
      }
      seenEnsureSheet |= isEnsureSheet(step);
    }
    if (!seenEnsureSheet) {
      return Optional.of(STREAMING_WRITE.missingEnsureSheetMutationMessage());
    }
    return Optional.empty();
  }

  private static Optional<String> streamingWriteStepFailure(
      WorkbookStep step, boolean seenEnsureSheet) {
    return switch (step) {
      case MutationStep mutationStep ->
          unsupportedStreamingMutationAction(mutationStep.action())
              .or(
                  () ->
                      requiresEnsureSheetBeforeAppend(mutationStep.action(), seenEnsureSheet)
                          ? Optional.of(STREAMING_WRITE.missingEnsureSheetBeforeAppendMessage())
                          : Optional.empty());
      case AssertionStep _ ->
          missingEnsureSheetBeforeObservation(
              seenEnsureSheet, STREAMING_WRITE.missingEnsureSheetBeforeAssertionMessage());
      case InspectionStep _ ->
          missingEnsureSheetBeforeObservation(
              seenEnsureSheet, STREAMING_WRITE.missingEnsureSheetBeforeInspectionMessage());
    };
  }

  private static Optional<String> unsupportedStreamingMutationAction(MutationAction action) {
    return STREAMING_WRITE_MUTATION_ACTION_TYPES.contains(action.getClass())
        ? Optional.empty()
        : Optional.of(STREAMING_WRITE.unsupportedActionMessage(action.actionType()));
  }

  private static boolean requiresEnsureSheetBeforeAppend(
      MutationAction action, boolean seenEnsureSheet) {
    return action instanceof CellMutationAction.AppendRow && !seenEnsureSheet;
  }

  private static Optional<String> missingEnsureSheetBeforeObservation(
      boolean seenEnsureSheet, String message) {
    return seenEnsureSheet ? Optional.empty() : Optional.of(message);
  }

  private static boolean isEnsureSheet(WorkbookStep step) {
    return step instanceof MutationStep mutationStep
        && mutationStep.action() instanceof WorkbookMutationAction.EnsureSheet;
  }
}
