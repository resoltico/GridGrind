package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.nio.file.Path;
import java.time.Instant;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;

/** Captures one structured execution journal while optionally streaming verbose events. */
final class ExecutionJournalRecorder {
  private final @Nullable String planId;
  private final ExecutionJournalLevel level;
  private final ExecutionJournal.SourceSummary source;
  private final ExecutionJournalSink sink;
  private final long planStartNanos;
  private final List<ExecutionJournal.Step> steps = new ArrayList<>();
  private final List<ExecutionJournal.Event> events = new ArrayList<>();
  private ExecutionJournal.Phase validation = ExecutionJournal.Phase.notStarted();
  private ExecutionJournal.Phase inputResolution = ExecutionJournal.Phase.notStarted();
  private ExecutionJournal.Phase open = ExecutionJournal.Phase.notStarted();
  private ExecutionJournal.Calculation calculation =
      new ExecutionJournal.Calculation(
          ExecutionJournal.Phase.notStarted(), ExecutionJournal.Phase.notStarted());
  private ExecutionJournal.Phase persistencePhase = ExecutionJournal.Phase.notStarted();
  private ExecutionJournal.Phase close = ExecutionJournal.Phase.notStarted();

  private ExecutionJournalRecorder(
      @Nullable String planId,
      ExecutionJournalLevel level,
      ExecutionJournal.SourceSummary source,
      ExecutionJournalSink sink) {
    this.planId = planId;
    this.level = level;
    this.source = source;
    this.sink = sink;
    this.planStartNanos = System.nanoTime();
    emit("PLAN", "started", null, null);
  }

  static ExecutionJournalRecorder start(
      WorkbookPlan request, ExecutionJournalSink sink, Path workingDirectory) {
    ExecutionJournalSink liveSink = ExecutionJournalSink.requireNonNull(sink);
    String planId = request == null ? null : request.planId().orElse(null);
    ExecutionJournalLevel level =
        request == null ? ExecutionJournalLevel.SUMMARY : request.journalLevel();
    ExecutionJournal.SourceSummary source =
        request == null
            ? new ExecutionJournal.SourceSummary(Optional.empty(), Optional.empty())
            : new ExecutionJournal.SourceSummary(
                Optional.of(ExecutionRequestPaths.reqSourceType(request)),
                Optional.ofNullable(
                    ExecutionRequestPaths.reqSourcePath(request, workingDirectory)));
    return new ExecutionJournalRecorder(planId, level, source, liveSink);
  }

  PhaseHandle beginValidation() {
    return new PhaseHandle("VALIDATION", null, null, phase -> validation = phase);
  }

  PhaseHandle beginOpen() {
    return new PhaseHandle("OPEN", null, null, phase -> open = phase);
  }

  PhaseHandle beginInputResolution() {
    return new PhaseHandle("RESOLVE_INPUTS", null, null, phase -> inputResolution = phase);
  }

  PhaseHandle beginPersistence() {
    return new PhaseHandle("PERSIST", null, null, phase -> persistencePhase = phase);
  }

  PhaseHandle beginClose() {
    return new PhaseHandle("CLOSE", null, null, phase -> close = phase);
  }

  PhaseHandle beginCalculationPreflight() {
    return new PhaseHandle(
        "CALCULATION_PREFLIGHT",
        null,
        null,
        phase -> calculation = new ExecutionJournal.Calculation(phase, calculation.execution()));
  }

  PhaseHandle beginCalculationExecution() {
    return new PhaseHandle(
        "CALCULATION_EXECUTION",
        null,
        null,
        phase -> calculation = new ExecutionJournal.Calculation(calculation.preflight(), phase));
  }

  void markCalculationPreflightNotRequested() {
    calculation =
        new ExecutionJournal.Calculation(
            ExecutionJournal.Phase.notRequested(), calculation.execution());
  }

  void markCalculationExecutionNotRequested() {
    calculation =
        new ExecutionJournal.Calculation(
            calculation.preflight(), ExecutionJournal.Phase.notRequested());
  }

  StepHandle beginStep(int stepIndex, WorkbookStep step) {
    return new StepHandle(stepIndex, step);
  }

  ExecutionJournal buildSuccess(int plannedStepCount) {
    return buildSuccess(plannedStepCount, true);
  }

  ExecutionJournal buildSuccess(int plannedStepCount, boolean emitPlanOutcomeEvent) {
    if (emitPlanOutcomeEvent) {
      emit("PLAN", "succeeded", null, null);
    }
    return new ExecutionJournal(
        Optional.ofNullable(planId),
        level,
        source,
        validation,
        inputResolution,
        open,
        calculation,
        persistencePhase,
        close,
        journalSteps(),
        ExecutionJournal.Outcome.succeeded(
            plannedStepCount, completedStepCount(), outcomeDurationMillis()),
        level == ExecutionJournalLevel.VERBOSE ? List.copyOf(events) : List.of());
  }

  ExecutionJournal buildFailure(
      int plannedStepCount,
      GridGrindProblemCode failureCode,
      @Nullable Integer failedStepIndex,
      @Nullable String failedStepId) {
    return buildFailure(plannedStepCount, failureCode, failedStepIndex, failedStepId, true);
  }

  ExecutionJournal buildFailure(
      int plannedStepCount,
      GridGrindProblemCode failureCode,
      @Nullable Integer failedStepIndex,
      @Nullable String failedStepId,
      boolean emitPlanOutcomeEvent) {
    if (emitPlanOutcomeEvent) {
      emit("PLAN", "failed (" + failureCode + ")", failedStepIndex, failedStepId);
    }
    return new ExecutionJournal(
        Optional.ofNullable(planId),
        level,
        source,
        validation,
        inputResolution,
        open,
        calculation,
        persistencePhase,
        close,
        journalSteps(),
        ExecutionJournal.Outcome.failed(
            plannedStepCount,
            completedStepCount(),
            outcomeDurationMillis(),
            Objects.requireNonNull(failureCode, "failureCode must not be null"),
            failedStepIndex == null || failedStepId == null
                ? Optional.empty()
                : Optional.of(new ExecutionJournal.FailureStep(failedStepIndex, failedStepId))),
        level == ExecutionJournalLevel.VERBOSE ? List.copyOf(events) : List.of());
  }

  private List<ExecutionJournal.Step> journalSteps() {
    return switch (level) {
      case SUMMARY ->
          steps.stream()
              .map(
                  step ->
                      new ExecutionJournal.Step(
                          step.stepIndex(),
                          step.stepId(),
                          step.stepKind(),
                          step.stepType(),
                          step.resolvedTargets(),
                          phaseWithoutTiming(step.phase()),
                          step.outcome(),
                          step.failure()))
              .toList();
      case NORMAL, VERBOSE -> List.copyOf(steps);
    };
  }

  static ExecutionJournal.Phase phaseWithoutTiming(ExecutionJournal.Phase phase) {
    return switch (phase) {
      case ExecutionJournal.Phase.NotStarted notStarted -> notStarted;
      case ExecutionJournal.Phase.NotRequested notRequested -> notRequested;
      case ExecutionJournal.Phase.Succeeded _ -> ExecutionJournal.Phase.succeededWithoutTiming();
      case ExecutionJournal.Phase.Failed _ -> ExecutionJournal.Phase.failedWithoutTiming();
    };
  }

  private int completedStepCount() {
    return (int)
        steps.stream()
            .filter(step -> step.outcome() == ExecutionJournal.StepOutcome.SUCCEEDED)
            .count();
  }

  private void emit(
      String category, String detail, @Nullable Integer stepIndex, @Nullable String stepId) {
    if (level != ExecutionJournalLevel.VERBOSE) {
      return;
    }
    ExecutionJournal.Event event =
        new ExecutionJournal.Event(
            Instant.now().toString(),
            category,
            detail,
            Optional.ofNullable(stepIndex),
            Optional.ofNullable(stepId));
    events.add(event);
    sink.emit(event);
  }

  private static long elapsedMillis(long startedAtNanos) {
    return (System.nanoTime() - startedAtNanos) / 1_000_000L;
  }

  private boolean recordsTiming() {
    return level != ExecutionJournalLevel.SUMMARY;
  }

  private long outcomeDurationMillis() {
    return recordsTiming() ? elapsedMillis(planStartNanos) : 0L;
  }

  /** Mutable handle that completes one top-level or nested execution phase exactly once. */
  final class PhaseHandle {
    private final String category;
    private final @Nullable Integer stepIndex;
    private final @Nullable String stepId;
    private final java.util.function.Consumer<ExecutionJournal.Phase> consumer;
    private final @Nullable String startedAt;
    private final long startedAtNanos;
    private boolean finished;

    private PhaseHandle(
        String category,
        @Nullable Integer stepIndex,
        @Nullable String stepId,
        java.util.function.Consumer<ExecutionJournal.Phase> consumer) {
      this.category = category;
      this.stepIndex = stepIndex;
      this.stepId = stepId;
      this.consumer = consumer;
      this.startedAt = recordsTiming() ? Instant.now().toString() : null;
      this.startedAtNanos = recordsTiming() ? System.nanoTime() : 0L;
      emit(category, "started", stepIndex, stepId);
    }

    ExecutionJournal.Phase succeed() {
      return finish("succeeded", true);
    }

    ExecutionJournal.Phase fail(String detail) {
      return finish(detail, false);
    }

    private ExecutionJournal.Phase finish(String detail, boolean succeeded) {
      if (finished) {
        throw new IllegalStateException("phase already finished: " + category);
      }
      finished = true;
      ExecutionJournal.Phase phase = finishedPhase(succeeded);
      consumer.accept(phase);
      emit(category, detail, stepIndex, stepId);
      return phase;
    }

    private ExecutionJournal.Phase finishedPhase(boolean succeeded) {
      if (!recordsTiming()) {
        return succeeded
            ? ExecutionJournal.Phase.succeededWithoutTiming()
            : ExecutionJournal.Phase.failedWithoutTiming();
      }
      String phaseStartedAt = Objects.requireNonNull(startedAt, "startedAt must not be null");
      String finishedAt = Instant.now().toString();
      long durationMillis = elapsedMillis(startedAtNanos);
      return succeeded
          ? ExecutionJournal.Phase.succeeded(phaseStartedAt, finishedAt, durationMillis)
          : ExecutionJournal.Phase.failed(phaseStartedAt, finishedAt, durationMillis);
    }
  }

  /** Mutable handle that collects timing, outcome, and calculation telemetry for one step. */
  final class StepHandle {
    private final int stepIndex;
    private final WorkbookStep step;
    private final PhaseHandle phaseHandle;

    private StepHandle(int stepIndex, WorkbookStep step) {
      this.stepIndex = stepIndex;
      this.step = step;
      this.phaseHandle = new PhaseHandle("STEP", stepIndex, step.stepId(), phase -> {});
    }

    void succeed() {
      ExecutionJournal.Phase phase = phaseHandle.succeed();
      steps.add(
          new ExecutionJournal.Step(
              stepIndex,
              step.stepId(),
              step.stepKind(),
              ExecutionStepKinds.stepType(step),
              ExecutionJournalTargetResolver.resolve(step, level),
              phase,
              ExecutionJournal.StepOutcome.SUCCEEDED,
              Optional.empty()));
    }

    void fail(
        GridGrindProblemCode code,
        dev.erst.gridgrind.contract.dto.GridGrindProblemCategory category,
        String stage,
        String message) {
      ExecutionJournal.Phase phase = phaseHandle.fail("failed (" + code + ")");
      steps.add(
          new ExecutionJournal.Step(
              stepIndex,
              step.stepId(),
              step.stepKind(),
              ExecutionStepKinds.stepType(step),
              ExecutionJournalTargetResolver.resolve(step, level),
              phase,
              ExecutionJournal.StepOutcome.FAILED,
              Optional.of(
                  new ExecutionJournal.FailureClassification(code, category, stage, message))));
    }
  }
}
