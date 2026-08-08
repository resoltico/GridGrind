package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookTempFileFactory;
import java.io.IOException;
import java.nio.file.Path;

/** Shared explicit-root helpers for executor tests after ambient execution removal. */
final class ExecutionContextFixtureSupport {
  private ExecutionContextFixtureSupport() {}

  static Path defaultWorkingDirectory() {
    return Path.of(System.getProperty("java.io.tmpdir")).toAbsolutePath().normalize();
  }

  static ExecutionInputBindings defaultBindings() {
    return ExecutionInputBindingsFixtureSupport.bindings(defaultWorkingDirectory());
  }

  static WorkbookTempFileFactory tempFileFactory(Path workingDirectory) {
    return WorkbookTempFileFactory.rooted(
        workingDirectory.toAbsolutePath().normalize().resolve(".gridgrind").resolve("tmp"));
  }

  static WorkbookTempFileFactory tempFileFactoryFor(Path anchoredPath) {
    Path normalized = anchoredPath.toAbsolutePath().normalize();
    Path parent = normalized.getParent();
    return tempFileFactory(parent == null ? normalized : parent);
  }

  static ExecutionWorkbookSupport workbookSupport(Path workingDirectory) {
    return new ExecutionWorkbookSupport(tempFileFactory(workingDirectory)::createTempFile);
  }

  static ExecutionWorkbookSupport defaultWorkbookSupport() {
    return workbookSupport(defaultWorkingDirectory());
  }

  static ExcelWorkbook openWorkbook(Path workbookPath) throws IOException {
    return ExcelWorkbooks.open(workbookPath, tempFileFactoryFor(workbookPath));
  }

  static ExecutionJournalRecorder startJournal(WorkbookPlan request, ExecutionJournalSink sink) {
    return ExecutionJournalRecorder.start(request, sink, defaultWorkingDirectory());
  }

  static ExecutionJournalRecorder startJournal(WorkbookPlan request) {
    return startJournal(request, ExecutionJournalSink.NOOP);
  }

  static ExecutionJournalRecorder startJournal(
      WorkbookPlan request, ExecutionJournalSink sink, Path workingDirectory) {
    return ExecutionJournalRecorder.start(request, sink, workingDirectory);
  }

  static WorkbookResult execute(DefaultGridGrindRequestExecutor executor, WorkbookPlan request) {
    return executor.execute(request, defaultBindings());
  }

  static WorkbookResult execute(
      DefaultGridGrindRequestExecutor executor, WorkbookPlan request, Path workingDirectory) {
    return executor.execute(
        request, ExecutionInputBindingsFixtureSupport.bindings(workingDirectory));
  }

  static void saveWorkbook(ExcelWorkbook workbook, Path workbookPath) throws IOException {
    workbook
        .persistence()
        .save(
            workbookPath,
            dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
            tempFileFactoryFor(workbookPath));
  }
}
