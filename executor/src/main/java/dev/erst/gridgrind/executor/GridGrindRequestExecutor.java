package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.Objects;

/** Transport-neutral execution port for one complete GridGrind request workflow. */
@FunctionalInterface
public interface GridGrindRequestExecutor {
  /** Executes the request and returns the corresponding structured response. */
  GridGrindResponse execute(
      WorkbookPlan request, ExecutionInputBindings bindings, ExecutionJournalSink sink);

  /** Executes the request with explicit authored-input bindings and no live journal sink. */
  default GridGrindResponse execute(WorkbookPlan request, ExecutionInputBindings bindings) {
    Objects.requireNonNull(bindings, "bindings must not be null");
    return execute(request, bindings, ExecutionJournalSink.NOOP);
  }

  /** Executes the request while optionally streaming verbose journal events to the sink. */
  default GridGrindResponse execute(WorkbookPlan request, ExecutionJournalSink sink) {
    Objects.requireNonNull(sink, "sink must not be null");
    return execute(request, ExecutionInputBindings.processDefault(), sink);
  }

  /** Executes the request with default process bindings and no live journal sink. */
  default GridGrindResponse execute(WorkbookPlan request) {
    return execute(request, ExecutionInputBindings.processDefault(), ExecutionJournalSink.NOOP);
  }

  /** Returns an executor that rejects null delegates up front. */
  static GridGrindRequestExecutor requireNonNull(GridGrindRequestExecutor executor) {
    return Objects.requireNonNull(executor, "executor must not be null");
  }
}
