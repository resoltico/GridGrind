package dev.erst.gridgrind.engine.api;

import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.Objects;

/** Transport-neutral execution port for one complete GridGrind request workflow. */
@FunctionalInterface
public interface GridGrindRequestExecutor {
  /** Executes the request and returns the corresponding structured response. */
  WorkbookResult execute(
      WorkbookPlan request, GridGrindRequestInputs inputs, GridGrindJournalSink sink);

  /** Executes the request with explicit authored-input bindings and no live journal sink. */
  default WorkbookResult execute(WorkbookPlan request, GridGrindRequestInputs inputs) {
    Objects.requireNonNull(inputs, "inputs must not be null");
    return execute(request, inputs, GridGrindJournalSink.NOOP);
  }

  /** Returns an executor that rejects null delegates up front. */
  static GridGrindRequestExecutor requireNonNull(GridGrindRequestExecutor executor) {
    return Objects.requireNonNull(executor, "executor must not be null");
  }
}
