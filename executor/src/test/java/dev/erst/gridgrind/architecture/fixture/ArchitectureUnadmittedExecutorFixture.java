package dev.erst.gridgrind.architecture.fixture;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.engine.runtime.ExecutionInputBindings;
import dev.erst.gridgrind.engine.runtime.ExecutionProgressSink;
import dev.erst.gridgrind.engine.runtime.GridGrindRequestExecutor;

/** Deliberately bypassing executor used only to prove the production architecture rule fails. */
public final class ArchitectureUnadmittedExecutorFixture implements GridGrindRequestExecutor {
  @Override
  public WorkbookResult execute(
      WorkbookPlan request, ExecutionInputBindings bindings, ExecutionProgressSink sink) {
    throw new UnsupportedOperationException("fixture must never execute workbook work");
  }
}
