package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.List;
import java.util.Objects;

/** Structured execution telemetry returned for every GridGrind run, including failures. */
public record ExecutionJournal(
    String planId,
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
    List<RequestWarning> warnings,
    Outcome outcome,
    @JsonInclude(JsonInclude.Include.NON_EMPTY) List<Event> events) {
  public ExecutionJournal {
    WorkbookPlan.requireNonBlank(planId, "planId");
    level = level == null ? ExecutionJournalLevel.NORMAL : level;
    Objects.requireNonNull(source, "source must not be null");
    Objects.requireNonNull(persistence, "persistence must not be null");
    Objects.requireNonNull(validation, "validation must not be null");
    Objects.requireNonNull(inputResolution, "inputResolution must not be null");
    Objects.requireNonNull(open, "open must not be null");
    Objects.requireNonNull(calculation, "calculation must not be null");
    Objects.requireNonNull(persistencePhase, "persistencePhase must not be null");
    Objects.requireNonNull(close, "close must not be null");
    steps = copyValues(steps, "steps");
    warnings = warnings == null ? List.of() : copyValues(warnings, "warnings");
    Objects.requireNonNull(outcome, "outcome must not be null");
    events = events == null ? List.of() : copyValues(events, "events");
  }

  /** Summary of the authored workbook source for one execution. */
  public record SourceSummary(String type, @JsonInclude(JsonInclude.Include.NON_NULL) String path) {
    public SourceSummary {
      WorkbookPlan.requireNonBlank(type, "type");
    }
  }

  /** Summary of the authored persistence policy for one execution. */
  public record PersistenceSummary(
      String type, @JsonInclude(JsonInclude.Include.NON_NULL) String path) {
    public PersistenceSummary {
      WorkbookPlan.requireNonBlank(type, "type");
    }
  }

  /** One timed execution phase. */
  public record Phase(
      Status status,
      @JsonInclude(JsonInclude.Include.NON_NULL) String startedAt,
      @JsonInclude(JsonInclude.Include.NON_NULL) String finishedAt,
      long durationMillis) {
    public Phase {
      Objects.requireNonNull(status, "status must not be null");
      if (status == Status.NOT_STARTED) {
        if (startedAt != null || finishedAt != null || durationMillis != 0) {
          throw new IllegalArgumentException(
              "NOT_STARTED phases must omit timestamps and use durationMillis=0");
        }
      } else {
        WorkbookPlan.requireNonBlank(startedAt, "startedAt");
        WorkbookPlan.requireNonBlank(finishedAt, "finishedAt");
        if (durationMillis < 0) {
          throw new IllegalArgumentException("durationMillis must be >= 0");
        }
      }
    }

    /** Creates a not-started phase. */
    public static Phase notStarted() {
      return new Phase(Status.NOT_STARTED, null, null, 0);
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
      @JsonInclude(JsonInclude.Include.NON_NULL) FailureClassification failure) {
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
      if (outcome == StepOutcome.FAILED && failure == null) {
        throw new IllegalArgumentException("failure must be present when outcome is FAILED");
      }
      if (outcome != StepOutcome.FAILED && failure != null) {
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
  public record Outcome(
      Status status,
      int plannedStepCount,
      int completedStepCount,
      long durationMillis,
      @JsonInclude(JsonInclude.Include.NON_NULL) Integer failedStepIndex,
      @JsonInclude(JsonInclude.Include.NON_NULL) String failedStepId,
      @JsonInclude(JsonInclude.Include.NON_NULL) GridGrindProblemCode failureCode) {
    public Outcome {
      Objects.requireNonNull(status, "status must not be null");
      if (plannedStepCount < 0) {
        throw new IllegalArgumentException("plannedStepCount must be >= 0");
      }
      if (completedStepCount < 0 || completedStepCount > plannedStepCount) {
        throw new IllegalArgumentException(
            "completedStepCount must be >= 0 and <= plannedStepCount");
      }
      if (durationMillis < 0) {
        throw new IllegalArgumentException("durationMillis must be >= 0");
      }
      if (status == Status.FAILED && failureCode == null) {
        throw new IllegalArgumentException("failureCode must be present when status is FAILED");
      }
      if (status != Status.FAILED && (failedStepIndex != null || failedStepId != null)) {
        throw new IllegalArgumentException(
            "failedStepIndex and failedStepId are only permitted when status is FAILED");
      }
      if (status != Status.FAILED && failureCode != null) {
        throw new IllegalArgumentException("failureCode is only permitted when status is FAILED");
      }
      if (failedStepId != null) {
        WorkbookPlan.requireNonBlank(failedStepId, "failedStepId");
      }
    }
  }

  /** Fine-grained event emitted for verbose journals and CLI live rendering. */
  public record Event(
      String timestamp,
      String category,
      String detail,
      @JsonInclude(JsonInclude.Include.NON_NULL) Integer stepIndex,
      @JsonInclude(JsonInclude.Include.NON_NULL) String stepId) {
    public Event {
      WorkbookPlan.requireNonBlank(timestamp, "timestamp");
      WorkbookPlan.requireNonBlank(category, "category");
      WorkbookPlan.requireNonBlank(detail, "detail");
      if (stepId != null) {
        WorkbookPlan.requireNonBlank(stepId, "stepId");
      }
    }
  }

  /** Status model shared by top-level and per-step phases. */
  public enum Status {
    NOT_STARTED,
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
}
