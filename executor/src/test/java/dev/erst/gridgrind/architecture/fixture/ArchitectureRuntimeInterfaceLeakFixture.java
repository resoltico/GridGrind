package dev.erst.gridgrind.architecture.fixture;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.engine.runtime.ExecutionInputBindings;
import dev.erst.gridgrind.engine.runtime.ExecutionProgressSink;
import dev.erst.gridgrind.engine.runtime.GridGrindRequestExecutor;

/** Deliberately invalid public type that inherits an unexported runtime interface. */
public final class ArchitectureRuntimeInterfaceLeakFixture implements GridGrindRequestExecutor {
  @Override
  public WorkbookResult execute(
      WorkbookPlan request, ExecutionInputBindings bindings, ExecutionProgressSink sink) {
    throw new UnsupportedOperationException("fixture method must not execute");
  }
}
