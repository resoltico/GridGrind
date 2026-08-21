package dev.erst.gridgrind.engine.api;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.json.RequestAnalysis;
import java.util.Objects;

/** Transport-neutral execution port for one complete GridGrind request workflow. */
@FunctionalInterface
public interface GridGrindRequestExecutor {
  /** Executes the request and returns the corresponding structured response. */
  WorkbookResult execute(
      WorkbookPlan request, GridGrindRequestInputs inputs, GridGrindProgressSink sink);

  /** Executes the request with explicit authored-input bindings and no live journal sink. */
  default WorkbookResult execute(WorkbookPlan request, GridGrindRequestInputs inputs) {
    Objects.requireNonNull(inputs, "inputs must not be null");
    return execute(request, inputs, GridGrindProgressSink.NOOP);
  }

  /**
   * Executes a fully bound raw-request analysis while retaining authored locations for preflight
   * diagnostics.
   */
  default WorkbookResult execute(
      RequestAnalysis analysis, GridGrindRequestInputs inputs, GridGrindProgressSink sink) {
    Objects.requireNonNull(analysis, "analysis must not be null");
    Objects.requireNonNull(inputs, "inputs must not be null");
    Objects.requireNonNull(sink, "sink must not be null");
    return execute(analysis.requireCompletePlan(), inputs, sink);
  }

  /** Executes one fully bound raw request with no live journal sink. */
  default WorkbookResult execute(RequestAnalysis analysis, GridGrindRequestInputs inputs) {
    return execute(analysis, inputs, GridGrindProgressSink.NOOP);
  }

  /** Returns an executor that rejects null delegates up front. */
  static GridGrindRequestExecutor requireNonNull(GridGrindRequestExecutor executor) {
    return Objects.requireNonNull(executor, "executor must not be null");
  }
}
