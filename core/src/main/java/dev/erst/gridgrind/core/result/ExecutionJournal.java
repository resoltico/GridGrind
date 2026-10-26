package dev.erst.gridgrind.core.result;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCategory;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Structured execution telemetry returned for every GridGrind run, including failures. */
@CatalogType(
    id = "ExecutionJournal",
    summary =
        "Structured execution telemetry returned on every success and failure response, including validation, input resolution, open, calculation, step, persistence, and close phases.",
    displayOrder = 0)
public record ExecutionJournal(
    ExecutionJournalLevel level,
    SourceSummary source,
    Phase validation,
    Phase inputResolution,
    Phase open,
    Calculation calculation,
    Phase persistencePhase,
    Phase close,
    List<Step> steps,
    Outcome outcome) {
  public ExecutionJournal {
    level = Objects.requireNonNullElse(level, ExecutionJournalLevel.SUMMARY);
    Objects.requireNonNull(source, "source must not be null");
    Objects.requireNonNull(validation, "validation must not be null");
    Objects.requireNonNull(inputResolution, "inputResolution must not be null");
    Objects.requireNonNull(open, "open must not be null");
    Objects.requireNonNull(calculation, "calculation must not be null");
    Objects.requireNonNull(persistencePhase, "persistencePhase must not be null");
    Objects.requireNonNull(close, "close must not be null");
    steps = copyValues(steps, "steps");
    Objects.requireNonNull(outcome, "outcome must not be null");
  }

  /** Summary of the authored workbook source for one execution. */
  @CatalogType(
      id = "ExecutionJournalSourceSummary",
      summary = "Journal summary of the authored workbook source.",
      displayOrder = 1)
  public record SourceSummary(Optional<String> type, Optional<String> path) {
    public SourceSummary {
      type = normalizeOptional(type, "type");
      path = normalizeOptional(path, "path");
      if (path.isPresent() && type.isEmpty()) {
        throw new IllegalArgumentException("type must be present when path is present");
      }
    }
  }

  /** One timed execution phase. */
  @Union(discriminator = "status", catalog = true, displayOrder = 41)
  public sealed interface Phase
      permits dev.erst.gridgrind.core.result.ExecutionJournal.Phase.NotStarted,
          dev.erst.gridgrind.core.result.ExecutionJournal.Phase.NotRequested,
          dev.erst.gridgrind.core.result.ExecutionJournal.Phase.Succeeded,
          dev.erst.gridgrind.core.result.ExecutionJournal.Phase.Failed {

    /** Canonical phase status token. */
    Status status();

    /** Creates a not-started phase. */
    static Phase notStarted() {
      return new NotStarted();
    }

    /** Creates a not-requested phase. */
    static Phase notRequested() {
      return new NotRequested();
    }

    /** Creates a succeeded phase with fully specified timing. */
    static Phase succeeded(String startedAt, String finishedAt, long durationMillis) {
      return new Phase.Succeeded(Optional.of(new Timing(startedAt, finishedAt, durationMillis)));
    }

    /** Creates a failed phase with fully specified timing. */
    static Phase failed(String startedAt, String finishedAt, long durationMillis) {
      return new Phase.Failed(Optional.of(new Timing(startedAt, finishedAt, durationMillis)));
    }

    /** Creates a succeeded phase when summary-mode output omits timing. */
    static Phase succeededWithoutTiming() {
      return new Phase.Succeeded(Optional.empty());
    }

    /** Creates a failed phase when summary-mode output omits timing. */
    static Phase failedWithoutTiming() {
      return new Phase.Failed(Optional.empty());
    }

    /** Phase never began. */
    @Variant(
        id = "NOT_STARTED",
        summary = "Execution phase that never began and therefore exposes no timing payload.")
    record NotStarted() implements Phase {
      @Override
      public Status status() {
        return Status.NOT_STARTED;
      }
    }

    /** Phase was intentionally skipped. */
    @Variant(
        id = "NOT_REQUESTED",
        summary = "Execution phase intentionally skipped and therefore exposing no timing payload.")
    record NotRequested() implements Phase {
      @Override
      public Status status() {
        return Status.NOT_REQUESTED;
      }
    }

    /** Phase finished successfully with measured timing. */
    @Variant(
        id = "SUCCEEDED",
        summary = "Execution phase that completed successfully and may include one timing payload.")
    record Succeeded(Optional<Timing> timing) implements Phase {
      public Succeeded {
        timing = normalizeOptionalValue(timing, "timing");
      }

      @Override
      public Status status() {
        return Status.SUCCEEDED;
      }
    }

    /** Phase ended in failure with measured timing. */
    @Variant(
        id = "FAILED",
        summary = "Execution phase that ended in failure and may include one timing payload.")
    record Failed(Optional<Timing> timing) implements Phase {
      public Failed {
        timing = normalizeOptionalValue(timing, "timing");
      }

      @Override
      public Status status() {
        return Status.FAILED;
      }
    }
  }

  /** Coherent timing payload for one phase that actually ran. */
  @CatalogType(
      id = "ExecutionJournalTiming",
      summary = "Measured timestamps and duration for one execution phase that actually ran.",
      displayOrder = 6)
  public record Timing(String startedAt, String finishedAt, long durationMillis) {
    public Timing {
      startedAt = requireExecutionTimestamp(startedAt, "startedAt");
      finishedAt = requireExecutionTimestamp(finishedAt, "finishedAt");
      requireNonNegativeDuration(durationMillis);
    }
  }

  /** Per-step execution journal entry. */
  @CatalogType(
      id = "ExecutionJournalStep",
      summary =
          "Per-step execution telemetry with resolved targets, timing, outcome, and optional failure detail.",
      displayOrder = 2)
  public record Step(
      int stepIndex,
      StepId stepId,
      String stepKind,
      String stepType,
      List<Target> resolvedTargets,
      Phase phase,
      StepOutcome outcome,
      Optional<FailureClassification> failure) {
    public Step {
      if (stepIndex < 0) {
        throw new IllegalArgumentException("stepIndex must be >= 0");
      }
      java.util.Objects.requireNonNull(stepId, "stepId must not be null");
      WorkbookPlan.requireNonBlank(stepKind, "stepKind");
      WorkbookPlan.requireNonBlank(stepType, "stepType");
      resolvedTargets = copyValues(resolvedTargets, "resolvedTargets");
      Objects.requireNonNull(phase, "phase must not be null");
      Objects.requireNonNull(outcome, "outcome must not be null");
      failure = Objects.requireNonNullElseGet(failure, Optional::empty);
      if (outcome == StepOutcome.FAILED && failure.isEmpty()) {
        throw new IllegalArgumentException("failure must be present when outcome is FAILED");
      }
      if (outcome != StepOutcome.FAILED && failure.isPresent()) {
        throw new IllegalArgumentException("failure is only permitted when outcome is FAILED");
      }
    }
  }

  /** One canonical target entry recorded for a step journal. */
  @CatalogType(
      id = "ExecutionJournalTarget",
      summary = "One canonical target label recorded inside a step journal.",
      displayOrder = 3)
  public record Target(String kind, String label) {
    public Target {
      WorkbookPlan.requireNonBlank(kind, "kind");
      WorkbookPlan.requireNonBlank(label, "label");
    }
  }

  /** Structured failure classification for a failed step. */
  @CatalogType(
      id = "ExecutionJournalFailureClassification",
      summary = "Structured problem-code classification for one failed step.",
      displayOrder = 4)
  public record FailureClassification(
      GridGrindProblemCode code, GridGrindProblemCategory category, String stage, String message) {
    public FailureClassification {
      Objects.requireNonNull(code, "code must not be null");
      Objects.requireNonNull(category, "category must not be null");
      WorkbookPlan.requireNonBlank(stage, "stage");
      WorkbookPlan.requireNonBlank(message, "message");
    }
  }

  /** Top-level calculation phase timings recorded for one execution. */
  @CatalogType(
      id = "ExecutionJournalCalculation",
      summary = "Top-level calculation preflight and execution timings for one request.",
      displayOrder = 5)
  public record Calculation(Phase preflight, Phase execution) {
    public Calculation {
      if (preflight == null) {
        throw new IllegalArgumentException("preflight must not be null");
      }
      if (execution == null) {
        throw new IllegalArgumentException("execution must not be null");
      }
    }
  }

  /** Final execution outcome summary. */
  @Union(discriminator = "status", catalog = true, displayOrder = 42)
  public sealed interface Outcome
      permits dev.erst.gridgrind.core.result.ExecutionJournal.Outcome.Succeeded,
          dev.erst.gridgrind.core.result.ExecutionJournal.Outcome.Failed {

    /** Canonical execution outcome status token. */
    Status status();

    /** Creates a succeeded execution outcome. */
    static Outcome succeeded(int plannedStepCount, int completedStepCount, long durationMillis) {
      return new Outcome.Succeeded(plannedStepCount, completedStepCount, durationMillis);
    }

    /** Creates a failed execution outcome. */
    static Outcome failed(
        int plannedStepCount,
        int completedStepCount,
        long durationMillis,
        GridGrindProblemCode failureCode,
        Optional<FailureStep> failedStep) {
      return new Outcome.Failed(
          plannedStepCount, completedStepCount, durationMillis, failureCode, failedStep);
    }

    /** Execution finished successfully. */
    @Variant(id = "SUCCEEDED", summary = "Execution outcome summary for a successful run.")
    record Succeeded(int plannedStepCount, int completedStepCount, long durationMillis)
        implements Outcome {
      public Succeeded {
        validateExecutionOutcomeCounts(plannedStepCount, completedStepCount, durationMillis);
      }

      @Override
      public Status status() {
        return Status.SUCCEEDED;
      }
    }

    /** Execution failed with one canonical failing-step summary. */
    @Variant(
        id = "FAILED",
        summary =
            "Execution outcome summary for a failed run, including the canonical failure code and an optional failing-step reference.")
    record Failed(
        int plannedStepCount,
        int completedStepCount,
        long durationMillis,
        GridGrindProblemCode problemCode,
        Optional<FailureStep> failedStep)
        implements Outcome {
      public Failed {
        validateExecutionOutcomeCounts(plannedStepCount, completedStepCount, durationMillis);
        Objects.requireNonNull(problemCode, "failureCode must not be null");
        failedStep = normalizeOptionalValue(failedStep, "failedStep");
      }

      @Override
      public Status status() {
        return Status.FAILED;
      }
    }
  }

  /** Canonical failing-step reference when a failure is attributable to one authored step. */
  @CatalogType(
      id = "ExecutionJournalFailureStep",
      summary =
          "Canonical failing-step reference recorded when an execution failure is attributable to one authored step.",
      displayOrder = 7)
  public record FailureStep(int failedStepIndex, StepId failedStepId) {
    public FailureStep {
      if (failedStepIndex < 0) {
        throw new IllegalArgumentException("failedStepIndex must be >= 0");
      }
      Objects.requireNonNull(failedStepId, "failedStepId must not be null");
    }
  }

  /** Status model shared by top-level and per-step phases. */
  public enum Status {
    NOT_STARTED,
    NOT_REQUESTED,
    SUCCEEDED,
    FAILED
  }

  /** Step-specific outcome values. */
  public enum StepOutcome {
    SUCCEEDED,
    FAILED
  }

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new java.util.ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }

  private static Optional<String> normalizeOptional(Optional<String> value, String fieldName) {
    Optional<String> normalized = Objects.requireNonNullElseGet(value, Optional::empty);
    if (normalized.isEmpty()) {
      return Optional.empty();
    }
    return Optional.of(WorkbookPlan.requireNonBlank(normalized.orElseThrow(), fieldName));
  }

  private static <T> Optional<T> normalizeOptionalValue(Optional<T> value, String fieldName) {
    Optional<T> normalized = Objects.requireNonNullElseGet(value, Optional::empty);
    normalized.ifPresent(
        entry -> Objects.requireNonNull(entry, fieldName + " must not contain null"));
    return normalized;
  }

  private static String requireExecutionTimestamp(String value, String fieldName) {
    return WorkbookPlan.requireNonBlank(value, fieldName);
  }

  private static void requireNonNegativeDuration(long durationMillis) {
    if (durationMillis < 0) {
      throw new IllegalArgumentException("durationMillis must be >= 0");
    }
  }

  private static void validateExecutionOutcomeCounts(
      int plannedStepCount, int completedStepCount, long durationMillis) {
    if (plannedStepCount < 0) {
      throw new IllegalArgumentException("plannedStepCount must be >= 0");
    }
    if (completedStepCount < 0 || completedStepCount > plannedStepCount) {
      throw new IllegalArgumentException("completedStepCount must be >= 0 and <= plannedStepCount");
    }
    requireNonNegativeDuration(durationMillis);
  }
}
