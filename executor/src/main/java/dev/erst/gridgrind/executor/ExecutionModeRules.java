package dev.erst.gridgrind.executor;

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
      WorkbookPlan request, ExecutionModeSelection executionModes) {
    if (executionModes.readMode() == ExecutionModeInput.ReadMode.EVENT_READ) {
      Optional<String> eventReadFailure = eventReadFailure(request);
      if (eventReadFailure.isPresent()) {
        return eventReadFailure;
      }
    }
    if (executionModes.writeMode() == ExecutionModeInput.WriteMode.STREAMING_WRITE) {
      return streamingWriteFailure(request);
    }
    return Optional.empty();
  }

  static boolean directEventReadEligible(
      WorkbookPlan request, ExecutionModeSelection executionModes) {
    return executionModes.readMode() == ExecutionModeInput.ReadMode.EVENT_READ
        && executionModes.writeMode() == ExecutionModeInput.WriteMode.FULL_XSSF
        && CalculationPolicyExecutor.allowsEventRead(request.calculationPolicy())
        && request.steps().stream().allMatch(InspectionStep.class::isInstance)
        && request.persistence() instanceof WorkbookPlan.WorkbookPersistence.None
        && request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile;
  }

  static ExecutionModeSelection executionModes(WorkbookPlan request) {
    ExecutionModeInput executionMode = request.effectiveExecutionMode();
    return new ExecutionModeSelection(executionMode.readMode(), executionMode.writeMode());
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
      switch (step) {
        case MutationStep mutationStep -> {
          if (!STREAMING_WRITE_MUTATION_ACTION_TYPES.contains(mutationStep.action().getClass())) {
            return Optional.of(
                STREAMING_WRITE.unsupportedActionMessage(mutationStep.action().actionType()));
          }
          if (mutationStep.action() instanceof WorkbookMutationAction.EnsureSheet) {
            seenEnsureSheet = true;
          }
          if (mutationStep.action() instanceof CellMutationAction.AppendRow && !seenEnsureSheet) {
            return Optional.of(STREAMING_WRITE.missingEnsureSheetBeforeAppendMessage());
          }
        }
        case AssertionStep _ -> {
          if (!seenEnsureSheet) {
            return Optional.of(STREAMING_WRITE.missingEnsureSheetBeforeAssertionMessage());
          }
        }
        case InspectionStep _ -> {
          if (!seenEnsureSheet) {
            return Optional.of(STREAMING_WRITE.missingEnsureSheetBeforeInspectionMessage());
          }
        }
      }
    }
    if (!seenEnsureSheet) {
      return Optional.of(STREAMING_WRITE.missingEnsureSheetMutationMessage());
    }
    return Optional.empty();
  }
}
