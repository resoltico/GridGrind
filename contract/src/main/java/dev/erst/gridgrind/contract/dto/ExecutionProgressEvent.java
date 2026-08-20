package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;
import java.util.Optional;

/** One compact, secret-safe JSONL progress event emitted while a verbose request executes. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "status")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ExecutionProgressEvent.Started.class, name = "STARTED"),
  @JsonSubTypes.Type(value = ExecutionProgressEvent.Succeeded.class, name = "SUCCEEDED"),
  @JsonSubTypes.Type(value = ExecutionProgressEvent.Failed.class, name = "FAILED")
})
public sealed interface ExecutionProgressEvent
    permits ExecutionProgressEvent.Started,
        ExecutionProgressEvent.Succeeded,
        ExecutionProgressEvent.Failed {
  /** ISO-8601 instant at which the lifecycle transition occurred. */
  String timestamp();

  /** Stable lifecycle category for the transition. */
  Category category();

  /** Stable lifecycle state represented by this exact variant. */
  Status status();

  /** Optional authored step index, present exactly when {@link #stepId()} is present. */
  Optional<Integer> stepIndex();

  /** Optional authored step id, present exactly when {@link #stepIndex()} is present. */
  Optional<String> stepId();

  /** Creates a started lifecycle event. */
  static Started started(
      String timestamp, Category category, Optional<Integer> stepIndex, Optional<String> stepId) {
    return new Started(timestamp, category, stepIndex, stepId);
  }

  /** Creates a succeeded lifecycle event. */
  static Succeeded succeeded(
      String timestamp, Category category, Optional<Integer> stepIndex, Optional<String> stepId) {
    return new Succeeded(timestamp, category, stepIndex, stepId);
  }

  /** Creates a failed lifecycle event with the classified failure code. */
  static Failed failed(
      String timestamp,
      Category category,
      GridGrindProblemCode problemCode,
      Optional<Integer> stepIndex,
      Optional<String> stepId) {
    return new Failed(timestamp, category, problemCode, stepIndex, stepId);
  }

  /** A lifecycle transition that just began. */
  record Started(
      String timestamp,
      Category category,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Integer> stepIndex,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> stepId)
      implements ExecutionProgressEvent {
    public Started {
      validateCommon(timestamp, category, stepIndex, stepId);
      stepIndex = normalizeStepIndex(stepIndex);
      stepId = normalizeStepId(stepId);
    }

    @Override
    public Status status() {
      return Status.STARTED;
    }
  }

  /** A lifecycle transition that completed successfully. */
  record Succeeded(
      String timestamp,
      Category category,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Integer> stepIndex,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> stepId)
      implements ExecutionProgressEvent {
    public Succeeded {
      validateCommon(timestamp, category, stepIndex, stepId);
      stepIndex = normalizeStepIndex(stepIndex);
      stepId = normalizeStepId(stepId);
    }

    @Override
    public Status status() {
      return Status.SUCCEEDED;
    }
  }

  /** A lifecycle transition that failed with one classified problem code. */
  record Failed(
      String timestamp,
      Category category,
      GridGrindProblemCode problemCode,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Integer> stepIndex,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> stepId)
      implements ExecutionProgressEvent {
    public Failed {
      validateCommon(timestamp, category, stepIndex, stepId);
      Objects.requireNonNull(problemCode, "problemCode must not be null");
      stepIndex = normalizeStepIndex(stepIndex);
      stepId = normalizeStepId(stepId);
    }

    @Override
    public Status status() {
      return Status.FAILED;
    }
  }

  /** Stable lifecycle category for one execution progress event. */
  enum Category {
    PLAN,
    VALIDATION,
    RESOLVE_INPUTS,
    OPEN,
    CALCULATION_PREFLIGHT,
    CALCULATION_EXECUTION,
    STEP,
    PERSIST,
    CLOSE
  }

  /** Stable lifecycle state for one execution progress event. */
  enum Status {
    STARTED,
    SUCCEEDED,
    FAILED
  }

  private static void validateCommon(
      String timestamp, Category category, Optional<Integer> stepIndex, Optional<String> stepId) {
    WorkbookPlan.requireNonBlank(timestamp, "timestamp");
    Objects.requireNonNull(category, "category must not be null");
    Optional<Integer> normalizedIndex = normalizeStepIndex(stepIndex);
    Optional<String> normalizedId = normalizeStepId(stepId);
    if (normalizedIndex.isPresent() != normalizedId.isPresent()) {
      throw new IllegalArgumentException(
          "stepId and stepIndex must either both be present or both be absent");
    }
  }

  private static Optional<Integer> normalizeStepIndex(Optional<Integer> value) {
    Optional<Integer> normalized = Objects.requireNonNullElseGet(value, Optional::empty);
    normalized.ifPresent(
        index -> {
          if (index < 0) {
            throw new IllegalArgumentException("stepIndex must be >= 0");
          }
        });
    return normalized;
  }

  private static Optional<String> normalizeStepId(Optional<String> value) {
    Optional<String> normalized = Objects.requireNonNullElseGet(value, Optional::empty);
    return normalized.map(entry -> WorkbookPlan.requireNonBlank(entry, "stepId"));
  }
}
