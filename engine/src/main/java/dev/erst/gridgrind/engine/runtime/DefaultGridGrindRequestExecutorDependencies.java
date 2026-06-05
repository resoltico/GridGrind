package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookArtifactIo;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.stream.ExcelStreamingWorkbookWriter;
import java.util.Objects;

/** Owned executor seams for one DefaultGridGrindRequestExecutor instance. */
record DefaultGridGrindRequestExecutorDependencies(
    WorkbookExecutionEngine workbookEngine,
    WorkbookCloser workbookCloser,
    ReadableWorkbookCloser readableWorkbookCloser,
    StreamingCalculationApplier streamingCalculationApplier) {

  DefaultGridGrindRequestExecutorDependencies {
    Objects.requireNonNull(workbookEngine, "workbookEngine must not be null");
    Objects.requireNonNull(workbookCloser, "workbookCloser must not be null");
    Objects.requireNonNull(readableWorkbookCloser, "readableWorkbookCloser must not be null");
    Objects.requireNonNull(
        streamingCalculationApplier, "streamingCalculationApplier must not be null");
  }

  static DefaultGridGrindRequestExecutorDependencies production() {
    return new DefaultGridGrindRequestExecutorDependencies(
        new WorkbookExecutionEngine(),
        ExcelWorkbook::close,
        WorkbookArtifactIo.MaterializedWorkbook::close,
        ExcelStreamingWorkbookWriter::markRecalculateOnOpen);
  }
}
