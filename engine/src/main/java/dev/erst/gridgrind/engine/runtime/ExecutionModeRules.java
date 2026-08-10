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
import java.util.ArrayList;
import java.util.List;
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

  static List<String> calculationPolicyFailures(WorkbookPlan request) {
    return calculationPolicyFailures(request.calculationPolicy(), request.steps());
  }

  static List<String> calculationPolicyFailures(
      dev.erst.gridgrind.contract.dto.CalculationPolicyInput calculationPolicy,
      List<WorkbookStep> steps) {
    if (!CalculationPolicyExecutor.requiresMutationPrefix(calculationPolicy)) {
      return List.of();
    }
    boolean seenObservationStep = false;
    for (WorkbookStep step : steps) {
      if (step instanceof MutationStep) {
        if (seenObservationStep) {
          return List.of(
              "execution.calculation.strategy="
                  + calculationPolicy.effectiveStrategy().strategyType()
                  + " requires all MUTATION steps to appear before any ASSERTION or INSPECTION"
                  + " step so calculation can run once at the mutation-to-observation boundary");
        }
      } else {
        seenObservationStep = true;
      }
    }
    return List.of();
  }

  static List<String> executionModeFailures(
      WorkbookPlan request, ExecutionModeInput executionMode) {
    return executionModeFailures(
        executionMode,
        request.calculationPolicy(),
        java.util.Optional.of(request.source()),
        request.steps());
  }

  static List<String> executionModeFailures(
      ExecutionModeInput executionMode,
      dev.erst.gridgrind.contract.dto.CalculationPolicyInput calculationPolicy,
      java.util.Optional<WorkbookPlan.WorkbookSource> source,
      List<WorkbookStep> steps) {
    return switch (executionMode) {
      case ExecutionModeInput.FullXssf _ -> List.of();
      case ExecutionModeInput.EventRead _ -> eventReadFailures(calculationPolicy, steps);
      case ExecutionModeInput.StreamingWrite _ ->
          source
              .map(value -> streamingWriteFailures(calculationPolicy, value, steps))
              .orElseGet(List::of);
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

  private static List<String> eventReadFailures(
      dev.erst.gridgrind.contract.dto.CalculationPolicyInput calculationPolicy,
      List<WorkbookStep> steps) {
    List<String> failures = new ArrayList<>();
    if (!CalculationPolicyExecutor.allowsEventRead(calculationPolicy)) {
      failures.add(EVENT_READ.calculationFailureMessage());
    }
    for (WorkbookStep step : steps) {
      if (!(step instanceof InspectionStep inspectionStep)) {
        failures.add(EVENT_READ.unsupportedStepMessage(step.stepKind()));
        continue;
      }
      if (!EVENT_READ_INSPECTION_QUERY_TYPES.contains(inspectionStep.query().getClass())) {
        failures.add(EVENT_READ.unsupportedQueryMessage(inspectionStep.query().queryType()));
      }
    }
    return List.copyOf(failures);
  }

  private static List<String> streamingWriteFailures(
      dev.erst.gridgrind.contract.dto.CalculationPolicyInput calculationPolicy,
      WorkbookPlan.WorkbookSource source,
      List<WorkbookStep> steps) {
    List<String> failures = new ArrayList<>();
    if (!CalculationPolicyExecutor.allowsStreamingWrite(calculationPolicy)) {
      failures.add(STREAMING_WRITE.calculationFailureMessage());
    }
    if (!(source instanceof WorkbookPlan.WorkbookSource.New)) {
      failures.add(STREAMING_WRITE.invalidSourceMessage());
    }
    boolean seenEnsureSheet = false;
    for (WorkbookStep step : steps) {
      failures.addAll(streamingWriteStepFailures(step, seenEnsureSheet));
      seenEnsureSheet |= isEnsureSheet(step);
    }
    if (!seenEnsureSheet) {
      failures.add(STREAMING_WRITE.missingEnsureSheetMutationMessage());
    }
    return List.copyOf(failures);
  }

  private static List<String> streamingWriteStepFailures(
      WorkbookStep step, boolean seenEnsureSheet) {
    return switch (step) {
      case MutationStep mutationStep -> mutationFailures(mutationStep.action(), seenEnsureSheet);
      case AssertionStep _ ->
          missingEnsureSheetBeforeObservation(
              seenEnsureSheet, STREAMING_WRITE.missingEnsureSheetBeforeAssertionMessage());
      case InspectionStep _ ->
          missingEnsureSheetBeforeObservation(
              seenEnsureSheet, STREAMING_WRITE.missingEnsureSheetBeforeInspectionMessage());
    };
  }

  private static List<String> mutationFailures(MutationAction action, boolean seenEnsureSheet) {
    List<String> failures = new ArrayList<>();
    unsupportedStreamingMutationAction(action).ifPresent(failures::add);
    if (requiresEnsureSheetBeforeAppend(action, seenEnsureSheet)) {
      failures.add(STREAMING_WRITE.missingEnsureSheetBeforeAppendMessage());
    }
    return List.copyOf(failures);
  }

  private static java.util.Optional<String> unsupportedStreamingMutationAction(
      MutationAction action) {
    return STREAMING_WRITE_MUTATION_ACTION_TYPES.contains(action.getClass())
        ? java.util.Optional.empty()
        : java.util.Optional.of(STREAMING_WRITE.unsupportedActionMessage(action.actionType()));
  }

  private static boolean requiresEnsureSheetBeforeAppend(
      MutationAction action, boolean seenEnsureSheet) {
    return action instanceof CellMutationAction.AppendRow && !seenEnsureSheet;
  }

  private static List<String> missingEnsureSheetBeforeObservation(
      boolean seenEnsureSheet, String message) {
    return seenEnsureSheet ? List.of() : List.of(message);
  }

  private static boolean isEnsureSheet(WorkbookStep step) {
    return step instanceof MutationStep mutationStep
        && mutationStep.action() instanceof WorkbookMutationAction.EnsureSheet;
  }
}
