package dev.erst.gridgrind.contract.dto;

import java.util.Objects;

/** One non-fatal request warning produced during protocol execution. */
public record RequestWarning(int stepIndex, String stepId, String stepType, String message) {
  public RequestWarning {
    if (stepIndex < 0) {
      throw new IllegalArgumentException("stepIndex must not be negative");
    }
    Objects.requireNonNull(stepId, "stepId must not be null");
    if (stepId.isBlank()) {
      throw new IllegalArgumentException("stepId must not be blank");
    }
    Objects.requireNonNull(stepType, "stepType must not be null");
    if (stepType.isBlank()) {
      throw new IllegalArgumentException("stepType must not be blank");
    }
    Objects.requireNonNull(message, "message must not be null");
    if (message.isBlank()) {
      throw new IllegalArgumentException("message must not be blank");
    }
  }
}
