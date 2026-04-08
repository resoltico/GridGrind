package dev.erst.gridgrind.protocol.dto;

import java.util.Objects;

/** One non-fatal request warning produced during protocol execution. */
public record RequestWarning(int operationIndex, String operationType, String message) {
  public RequestWarning {
    if (operationIndex < 0) {
      throw new IllegalArgumentException("operationIndex must not be negative");
    }
    Objects.requireNonNull(operationType, "operationType must not be null");
    if (operationType.isBlank()) {
      throw new IllegalArgumentException("operationType must not be blank");
    }
    Objects.requireNonNull(message, "message must not be null");
    if (message.isBlank()) {
      throw new IllegalArgumentException("message must not be blank");
    }
  }
}
