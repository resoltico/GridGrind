package dev.erst.gridgrind.core.result;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.Objects;
import java.util.Optional;

/** One compact, secret-safe JSONL progress event emitted while a verbose request executes. */
@Union(discriminator = "status", catalog = true, displayOrder = 43)
public sealed interface ExecutionProgressEvent
    permits dev.erst.gridgrind.core.result.ExecutionProgressEvent.Started,
        dev.erst.gridgrind.core.result.ExecutionProgressEvent.Succeeded,
        dev.erst.gridgrind.core.result.ExecutionProgressEvent.Failed {
  /** ISO-8601 instant at which the lifecycle transition occurred. */
  String timestamp();

  /** Stable lifecycle category for the transition. */
  Category category();

  /** Stable lifecycle state represented by this exact variant. */
  Status status();

  /** Optional authored step index, present exactly when {@link #stepId()} is present. */
  Optional<Integer> stepIndex();

  /** Optional authored step id, present exactly when {@link #stepIndex()} is present. */
  Optional<StepId> stepId();

  /** Creates a started lifecycle event. */
  static Started started(
      String timestamp, Category category, Optional<Integer> stepIndex, Optional<StepId> stepId) {
    return new Started(timestamp, category, stepIndex, stepId);
  }

  /** Creates a succeeded lifecycle event. */
  static Succeeded succeeded(
      String timestamp, Category category, Optional<Integer> stepIndex, Optional<StepId> stepId) {
    return new Succeeded(timestamp, category, stepIndex, stepId);
  }

  /** Creates a failed lifecycle event with the classified failure code. */
  static Failed failed(
      String timestamp,
      Category category,
      GridGrindProblemCode problemCode,
      Optional<Integer> stepIndex,
      Optional<StepId> stepId) {
    return new Failed(timestamp, category, problemCode, stepIndex, stepId);
  }

  /** A lifecycle transition that just began. */
  @Variant(
      id = "STARTED",
      summary = "Compact JSONL lifecycle event emitted when one execution phase starts.")
  record Started(
      String timestamp, Category category, Optional<Integer> stepIndex, Optional<StepId> stepId)
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
  @Variant(
      id = "SUCCEEDED",
      summary = "Compact JSONL lifecycle event emitted when one execution phase succeeds.")
  record Succeeded(
      String timestamp, Category category, Optional<Integer> stepIndex, Optional<StepId> stepId)
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
  @Variant(
      id = "FAILED",
      summary =
          "Compact JSONL lifecycle event emitted when one execution phase fails with one classified problem code.")
  record Failed(
      String timestamp,
      Category category,
      GridGrindProblemCode problemCode,
      Optional<Integer> stepIndex,
      Optional<StepId> stepId)
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
      String timestamp, Category category, Optional<Integer> stepIndex, Optional<StepId> stepId) {
    WorkbookPlan.requireNonBlank(timestamp, "timestamp");
    Objects.requireNonNull(category, "category must not be null");
    Optional<Integer> normalizedIndex = normalizeStepIndex(stepIndex);
    Optional<StepId> normalizedId = normalizeStepId(stepId);
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

  private static Optional<StepId> normalizeStepId(Optional<StepId> value) {
    return Objects.requireNonNullElseGet(value, Optional::empty);
  }
}
