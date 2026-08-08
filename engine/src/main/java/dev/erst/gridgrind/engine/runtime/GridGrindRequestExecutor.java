package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import java.util.Objects;

/** Transport-neutral execution port for one complete GridGrind request workflow. */
@FunctionalInterface
public interface GridGrindRequestExecutor {
  /** Executes the request and returns the corresponding structured response. */
  WorkbookResult execute(
      WorkbookPlan request, ExecutionInputBindings bindings, ExecutionJournalSink sink);

  /** Executes the request with explicit authored-input bindings and no live journal sink. */
  default WorkbookResult execute(WorkbookPlan request, ExecutionInputBindings bindings) {
    Objects.requireNonNull(bindings, "bindings must not be null");
    return execute(request, bindings, ExecutionJournalSink.NOOP);
  }

  /** Returns an executor that rejects null delegates up front. */
  static GridGrindRequestExecutor requireNonNull(GridGrindRequestExecutor executor) {
    return Objects.requireNonNull(executor, "executor must not be null");
  }
}
