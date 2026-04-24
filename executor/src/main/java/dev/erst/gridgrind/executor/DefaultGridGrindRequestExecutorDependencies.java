package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.excel.ExcelOoxmlPackageSecuritySupport;
import dev.erst.gridgrind.excel.ExcelStreamingWorkbookWriter;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import java.nio.file.Files;
import java.util.Objects;

/** Owned executor seams for one DefaultGridGrindRequestExecutor instance. */
record DefaultGridGrindRequestExecutorDependencies(
    WorkbookCommandExecutor commandExecutor,
    WorkbookReadExecutor readExecutor,
    WorkbookCloser workbookCloser,
    TempFileFactory tempFileFactory,
    ReadableWorkbookCloser readableWorkbookCloser,
    StreamingCalculationApplier streamingCalculationApplier) {

  DefaultGridGrindRequestExecutorDependencies {
    Objects.requireNonNull(commandExecutor, "commandExecutor must not be null");
    Objects.requireNonNull(readExecutor, "readExecutor must not be null");
    Objects.requireNonNull(workbookCloser, "workbookCloser must not be null");
    Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
    Objects.requireNonNull(readableWorkbookCloser, "readableWorkbookCloser must not be null");
    Objects.requireNonNull(
        streamingCalculationApplier, "streamingCalculationApplier must not be null");
  }

  static DefaultGridGrindRequestExecutorDependencies production() {
    return new DefaultGridGrindRequestExecutorDependencies(
        new WorkbookCommandExecutor(),
        new WorkbookReadExecutor(),
        ExcelWorkbook::close,
        Files::createTempFile,
        ExcelOoxmlPackageSecuritySupport.ReadableWorkbook::close,
        ExcelStreamingWorkbookWriter::markRecalculateOnOpen);
  }
}
