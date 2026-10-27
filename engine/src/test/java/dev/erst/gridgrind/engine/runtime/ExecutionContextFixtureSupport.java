package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookTempFileFactory;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Objects;

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

  static WorkbookLifecycle workbookSupport(Path workingDirectory) {
    return new WorkbookLifecycle(tempFileFactory(workingDirectory)::createTempFile);
  }

  static WorkbookLifecycle defaultWorkbookSupport() {
    return workbookSupport(defaultWorkingDirectory());
  }

  static XssfWorkbook openWorkbook(Path workbookPath) throws IOException {
    return XssfWorkbooks.open(workbookPath, tempFileFactoryFor(workbookPath));
  }

  static ExecutionJournalRecorder startJournal(WorkbookPlan request, ExecutionProgressSink sink) {
    return ExecutionJournalRecorder.start(request, sink, defaultWorkingDirectory());
  }

  static ExecutionJournalRecorder startJournal(WorkbookPlan request) {
    return startJournal(request, ExecutionProgressSink.NOOP);
  }

  static ExecutionJournalRecorder startJournal(
      WorkbookPlan request, ExecutionProgressSink sink, Path workingDirectory) {
    return ExecutionJournalRecorder.start(request, sink, workingDirectory);
  }

  static WorkbookResult execute(DefaultGridGrindRequestExecutor executor, WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return executor.execute(
        request, ExecutionInputBindingsFixtureSupport.bindings(defaultWorkingDirectory(), request));
  }

  static WorkbookResult execute(
      DefaultGridGrindRequestExecutor executor, WorkbookPlan request, Path workingDirectory) {
    Objects.requireNonNull(request, "request must not be null");
    return executor.execute(
        request, ExecutionInputBindingsFixtureSupport.bindings(workingDirectory, request));
  }

  static void saveWorkbook(XssfWorkbook workbook, Path workbookPath) throws IOException {
    workbook
        .persistence()
        .save(
            workbookPath,
            dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                .REPLACE_EXISTING,
            tempFileFactoryFor(workbookPath));
  }
}
