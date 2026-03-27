package dev.erst.gridgrind.protocol;

import java.util.Objects;

/** Transport-neutral execution port for one complete GridGrind request workflow. */
@FunctionalInterface
public interface GridGrindRequestExecutor {
  /** Executes the request and returns the corresponding structured response. */
  GridGrindResponse execute(GridGrindRequest request);

  /** Returns an executor that rejects null delegates up front. */
  static GridGrindRequestExecutor requireNonNull(GridGrindRequestExecutor executor) {
    return Objects.requireNonNull(executor, "executor must not be null");
  }
}
