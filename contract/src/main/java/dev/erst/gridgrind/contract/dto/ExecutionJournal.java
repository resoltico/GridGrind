package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Structured execution telemetry returned for every GridGrind run, including failures. */
public record ExecutionJournal(
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> planId,
    ExecutionJournalLevel level,
    SourceSummary source,
    PersistenceSummary persistence,
    Phase validation,
    Phase inputResolution,
    Phase open,
    Calculation calculation,
    Phase persistencePhase,
    Phase close,
    List<Step> steps,
    Outcome outcome,
    @JsonInclude(JsonInclude.Include.NON_EMPTY) List<Event> events) {
  public ExecutionJournal {
    planId = Objects.requireNonNullElseGet(planId, Optional::empty);
    if (planId.isPresent()) {
      planId = Optional.of(WorkbookPlan.requireNonBlank(planId.orElseThrow(), "planId"));
    }
    level = Objects.requireNonNullElse(level, ExecutionJournalLevel.SUMMARY);
    Objects.requireNonNull(source, "source must not be null");
    Objects.requireNonNull(persistence, "persistence must not be null");
    Objects.requireNonNull(validation, "validation must not be null");
    Objects.requireNonNull(inputResolution, "inputResolution must not be null");
    Objects.requireNonNull(open, "open must not be null");
    Objects.requireNonNull(calculation, "calculation must not be null");
    Objects.requireNonNull(persistencePhase, "persistencePhase must not be null");
    Objects.requireNonNull(close, "close must not be null");
    steps = copyValues(steps, "steps");
    Objects.requireNonNull(outcome, "outcome must not be null");
    events = copyValues(Objects.requireNonNullElseGet(events, List::of), "events");
  }

  /** Summary of the authored workbook source for one execution. */
  public record SourceSummary(
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> type,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> path) {
    public SourceSummary {
      type = normalizeOptional(type, "type");
      path = normalizeOptional(path, "path");
      if (path.isPresent() && type.isEmpty()) {
        throw new IllegalArgumentException("type must be present when path is present");
      }
    }
  }

  /** Summary of the authored persistence policy for one execution. */
  public record PersistenceSummary(
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> type,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> path) {
    public PersistenceSummary {
      type = normalizeOptional(type, "type");
      path = normalizeOptional(path, "path");
      if (path.isPresent() && type.isEmpty()) {
        throw new IllegalArgumentException("type must be present when path is present");
      }
    }
  }

  /** One timed execution phase. */
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "status")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = Phase.NotStarted.class, name = "NOT_STARTED"),
    @JsonSubTypes.Type(value = Phase.NotRequested.class, name = "NOT_REQUESTED"),
    @JsonSubTypes.Type(value = Phase.Succeeded.class, name = "SUCCEEDED"),
    @JsonSubTypes.Type(value = Phase.Failed.class, name = "FAILED")
  })
  public sealed interface Phase
      permits Phase.NotStarted, Phase.NotRequested, Phase.Succeeded, Phase.Failed {

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
    record NotStarted() implements Phase {
      @Override
      public Status status() {
        return Status.NOT_STARTED;
      }
    }

    /** Phase was intentionally skipped. */
    record NotRequested() implements Phase {
      @Override
      public Status status() {
        return Status.NOT_REQUESTED;
      }
    }

    /** Phase finished successfully with measured timing. */
    record Succeeded(@JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Timing> timing)
        implements Phase {
      public Succeeded {
        timing = normalizeOptionalValue(timing, "timing");
      }

      @Override
      public Status status() {
        return Status.SUCCEEDED;
      }
    }

    /** Phase ended in failure with measured timing. */
    record Failed(@JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Timing> timing)
        implements Phase {
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
  public record Timing(String startedAt, String finishedAt, long durationMillis) {
    public Timing {
      startedAt = requireExecutionTimestamp(startedAt, "startedAt");
      finishedAt = requireExecutionTimestamp(finishedAt, "finishedAt");
      requireNonNegativeDuration(durationMillis);
    }
  }

  /** Per-step execution journal entry. */
  public record Step(
      int stepIndex,
      String stepId,
      String stepKind,
      String stepType,
      List<Target> resolvedTargets,
      Phase phase,
      StepOutcome outcome,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<FailureClassification> failure) {
    public Step {
      if (stepIndex < 0) {
        throw new IllegalArgumentException("stepIndex must be >= 0");
      }
      WorkbookPlan.requireNonBlank(stepId, "stepId");
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
  public record Target(String kind, String label) {
    public Target {
      WorkbookPlan.requireNonBlank(kind, "kind");
      WorkbookPlan.requireNonBlank(label, "label");
    }
  }

  /** Structured failure classification for a failed step. */
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
  @JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "status")
  @JsonSubTypes({
    @JsonSubTypes.Type(value = Outcome.Succeeded.class, name = "SUCCEEDED"),
    @JsonSubTypes.Type(value = Outcome.Failed.class, name = "FAILED")
  })
  public sealed interface Outcome permits Outcome.Succeeded, Outcome.Failed {

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
    record Failed(
        int plannedStepCount,
        int completedStepCount,
        long durationMillis,
        GridGrindProblemCode problemCode,
        @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<FailureStep> failedStep)
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
  public record FailureStep(int failedStepIndex, String failedStepId) {
    public FailureStep {
      if (failedStepIndex < 0) {
        throw new IllegalArgumentException("failedStepIndex must be >= 0");
      }
      WorkbookPlan.requireNonBlank(failedStepId, "failedStepId");
    }
  }

  /** Fine-grained event emitted for verbose journals and CLI live rendering. */
  public record Event(
      String timestamp,
      String category,
      String detail,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Integer> stepIndex,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> stepId) {
    public Event {
      WorkbookPlan.requireNonBlank(timestamp, "timestamp");
      WorkbookPlan.requireNonBlank(category, "category");
      WorkbookPlan.requireNonBlank(detail, "detail");
      stepIndex = Objects.requireNonNullElseGet(stepIndex, Optional::empty);
      stepId = normalizeOptional(stepId, "stepId");
      if (stepId.isPresent() != stepIndex.isPresent()) {
        throw new IllegalArgumentException(
            "stepId and stepIndex must either both be present or both be absent");
      }
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
