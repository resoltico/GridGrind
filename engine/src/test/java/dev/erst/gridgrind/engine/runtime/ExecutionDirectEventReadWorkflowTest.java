package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookTempFileFactory;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies event-read response truth when terminal host acceptance cannot read its workbook. */
class ExecutionDirectEventReadWorkflowTest {
  @TempDir Path root;

  @Test
  void returnsStructuredFailureWhenTerminalHostAcceptanceHasAnIoFailure() throws Exception {
    Path source = writeWorkbook();
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile(source.getFileName().toString()),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.mode(new ExecutionModeInput.EventRead()),
            FormulaEnvironmentInput.empty(),
            List.of());
    ExecutionInputBindings rawBindings =
        new ExecutionInputBindings(
            root,
            root.resolve("scratch"),
            ExecutionGrantTestSupport.forPlan(request, root, source));
    ExecutionDirectEventReadWorkflow workflow =
        new ExecutionDirectEventReadWorkflow(
            stepSupport(rawBindings),
            (policy, workbookPath, workbookLocation) -> {
              throw new IOException("terminal host acceptance storage unavailable");
            },
            responseSupport(),
            rawBindings.tempFileFactory());

    try (RequestPathAccess access =
        new RequestPathAccess(
            rawBindings.workingDirectory(),
            rawBindings.tempFileFactory(),
            rawBindings.executionGrant())) {
      WorkbookResult.Failure failure =
          assertInstanceOf(
              WorkbookResult.Failure.class,
              workflow.execute(
                  GridGrindProtocolVersion.current(),
                  request,
                  List.of(),
                  ExecutionJournalRecorder.start(request, ExecutionProgressSink.NOOP, root),
                  rawBindings.withRequestPathAccess(access)));

      assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    }
  }

  private ExecutionStepSupport stepSupport(ExecutionInputBindings bindings) {
    WorkbookExecutionEngine workbookEngine = new WorkbookExecutionEngine();
    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(workbookEngine);
    return new ExecutionStepSupport(
        workbookEngine,
        selectorResolver,
        new AssertionExecutor(workbookEngine, selectorResolver),
        WorkbookTempFileFactory.rooted(root.resolve("scratch"))::createTempFile,
        bindings);
  }

  private static ExecutionResponseSupport responseSupport() {
    return new ExecutionResponseSupport(ExcelWorkbook::close, materialized -> materialized.close());
  }

  private Path writeWorkbook() throws IOException {
    Path source = root.resolve("source.xlsx");
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        var output = Files.newOutputStream(source)) {
      workbook.createSheet("Budget");
      workbook.write(output);
    }
    return source;
  }
}
