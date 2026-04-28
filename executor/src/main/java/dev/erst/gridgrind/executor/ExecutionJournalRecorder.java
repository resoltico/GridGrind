package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import java.nio.file.Path;
import java.time.Instant;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.UUID;

/** Captures one structured execution journal while optionally streaming verbose events. */
final class ExecutionJournalRecorder {
  private final String planId;
  private final ExecutionJournalLevel level;
  private final ExecutionJournal.SourceSummary source;
  private final ExecutionJournal.PersistenceSummary persistence;
  private final ExecutionJournalSink sink;
  private final long planStartNanos;
  private final List<ExecutionJournal.Step> steps = new ArrayList<>();
  private final List<ExecutionJournal.Event> events = new ArrayList<>();
  private List<RequestWarning> warnings = List.of();
  private ExecutionJournal.Phase validation = ExecutionJournal.Phase.notStarted();
  private ExecutionJournal.Phase inputResolution = ExecutionJournal.Phase.notStarted();
  private ExecutionJournal.Phase open = ExecutionJournal.Phase.notStarted();
  private ExecutionJournal.Calculation calculation =
      new ExecutionJournal.Calculation(
          ExecutionJournal.Phase.notStarted(), ExecutionJournal.Phase.notStarted());
  private ExecutionJournal.Phase persistencePhase = ExecutionJournal.Phase.notStarted();
  private ExecutionJournal.Phase close = ExecutionJournal.Phase.notStarted();

  private ExecutionJournalRecorder(
      String planId,
      ExecutionJournalLevel level,
      ExecutionJournal.SourceSummary source,
      ExecutionJournal.PersistenceSummary persistence,
      ExecutionJournalSink sink) {
    this.planId = planId;
    this.level = level;
    this.source = source;
    this.persistence = persistence;
    this.sink = sink;
    this.planStartNanos = System.nanoTime();
    emit("PLAN", "started", null, null);
  }

  static ExecutionJournalRecorder start(WorkbookPlan request, ExecutionJournalSink sink) {
    return start(request, sink, Path.of(""));
  }

  static ExecutionJournalRecorder start(
      WorkbookPlan request, ExecutionJournalSink sink, Path workingDirectory) {
    ExecutionJournalSink liveSink = ExecutionJournalSink.requireNonNull(sink);
    String planId =
        request == null
            ? null
            : request
                .planId()
                .orElseGet(
                    () ->
                        "plan-" + UUID.randomUUID().toString().toLowerCase(java.util.Locale.ROOT));
    ExecutionJournalLevel level =
        request == null ? ExecutionJournalLevel.NORMAL : request.journalLevel();
    ExecutionJournal.SourceSummary source =
        request == null
            ? new ExecutionJournal.SourceSummary(Optional.empty(), Optional.empty())
            : new ExecutionJournal.SourceSummary(
                Optional.of(ExecutionRequestPaths.reqSourceType(request)),
                Optional.ofNullable(
                    ExecutionRequestPaths.reqSourcePath(request, workingDirectory)));
    ExecutionJournal.PersistenceSummary persistence =
        request == null
            ? new ExecutionJournal.PersistenceSummary(Optional.empty(), Optional.empty())
            : new ExecutionJournal.PersistenceSummary(
                Optional.of(ExecutionRequestPaths.reqPersistenceType(request)),
                Optional.ofNullable(
                    ExecutionRequestPaths.persistencePath(
                        request.source(), request.persistence(), workingDirectory)));
    return new ExecutionJournalRecorder(planId, level, source, persistence, liveSink);
  }

  void setWarnings(List<RequestWarning> warnings) {
    this.warnings = warnings == null ? List.of() : List.copyOf(warnings);
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

  StepHandle beginStep(int stepIndex, WorkbookStep step) {
    return new StepHandle(stepIndex, step);
  }

  ExecutionJournal buildSuccess(int plannedStepCount) {
    emit("PLAN", "succeeded", null, null);
    return new ExecutionJournal(
        Optional.ofNullable(planId),
        level,
        source,
        persistence,
        validation,
        inputResolution,
        open,
        calculation,
        persistencePhase,
        close,
        List.copyOf(steps),
        warnings,
        new ExecutionJournal.Outcome(
            ExecutionJournal.Status.SUCCEEDED,
            plannedStepCount,
            completedStepCount(),
            elapsedMillis(planStartNanos),
            Optional.empty(),
            Optional.empty(),
            Optional.empty()),
        level == ExecutionJournalLevel.VERBOSE ? List.copyOf(events) : List.of());
  }

  ExecutionJournal buildFailure(
      int plannedStepCount,
      GridGrindProblemCode failureCode,
      Integer failedStepIndex,
      String failedStepId) {
    emit("PLAN", "failed (" + failureCode + ")", failedStepIndex, failedStepId);
    return new ExecutionJournal(
        Optional.ofNullable(planId),
        level,
        source,
        persistence,
        validation,
        inputResolution,
        open,
        calculation,
        persistencePhase,
        close,
        List.copyOf(steps),
        warnings,
        new ExecutionJournal.Outcome(
            ExecutionJournal.Status.FAILED,
            plannedStepCount,
            completedStepCount(),
            elapsedMillis(planStartNanos),
            Optional.ofNullable(failedStepIndex),
            Optional.ofNullable(failedStepId),
            Optional.ofNullable(failureCode)),
        level == ExecutionJournalLevel.VERBOSE ? List.copyOf(events) : List.of());
  }

  private int completedStepCount() {
    return (int)
        steps.stream()
            .filter(step -> step.outcome() == ExecutionJournal.StepOutcome.SUCCEEDED)
            .count();
  }

  private void emit(String category, String detail, Integer stepIndex, String stepId) {
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

  /** Mutable handle that completes one top-level or nested execution phase exactly once. */
  final class PhaseHandle {
    private final String category;
    private final Integer stepIndex;
    private final String stepId;
    private final java.util.function.Consumer<ExecutionJournal.Phase> consumer;
    private final String startedAt;
    private final long startedAtNanos;
    private boolean finished;

    private PhaseHandle(
        String category,
        Integer stepIndex,
        String stepId,
        java.util.function.Consumer<ExecutionJournal.Phase> consumer) {
      this.category = category;
      this.stepIndex = stepIndex;
      this.stepId = stepId;
      this.consumer = consumer;
      this.startedAt = Instant.now().toString();
      this.startedAtNanos = System.nanoTime();
      emit(category, "started", stepIndex, stepId);
    }

    ExecutionJournal.Phase succeed() {
      return finish(ExecutionJournal.Status.SUCCEEDED, "succeeded");
    }

    ExecutionJournal.Phase fail(String detail) {
      return finish(ExecutionJournal.Status.FAILED, detail);
    }

    private ExecutionJournal.Phase finish(ExecutionJournal.Status status, String detail) {
      if (finished) {
        throw new IllegalStateException("phase already finished: " + category);
      }
      finished = true;
      ExecutionJournal.Phase phase =
          new ExecutionJournal.Phase(
              status,
              Optional.of(startedAt),
              Optional.of(Instant.now().toString()),
              elapsedMillis(startedAtNanos));
      consumer.accept(phase);
      emit(category, detail, stepIndex, stepId);
      return phase;
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
