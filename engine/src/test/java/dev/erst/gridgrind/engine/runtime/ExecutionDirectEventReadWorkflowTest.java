package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookTempFileFactory;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
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
            ExecutionPolicySpec.mode(new ExecutionModeSpec.EventRead()),
            FormulaEnvironmentSpec.empty(),
            List.of());
    ExecutionInputBindings rawBindings =
        new ExecutionInputBindings(
            root,
            root.resolve("scratch"),
            ExecutionGrantTestSupport.forPlan(request, root, source));
    ExecutionDirectEventReadWorkflow workflow =
        new ExecutionDirectEventReadWorkflow(
            stepSupport(rawBindings),
            (authoredPolicy, policy, workbookPath, workbookLocation, assertions) -> {
              throw new IOException("terminal host acceptance storage unavailable");
            },
            responseSupport());

    try (var prepared = RequestPreflight.verify(request, rawBindings).requirePreparedRequest()) {
      WorkbookResult.Failure failure =
          assertInstanceOf(
              WorkbookResult.Failure.class,
              workflow.execute(
                  new ExecutionDirectEventReadWorkflow.Call(
                      ProtocolVersion.CURRENT,
                      prepared,
                      List.of(),
                      ExecutionJournalRecorder.start(request, ExecutionProgressSink.NOOP, root))));

      assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    }
  }

  private StepExecution stepSupport(ExecutionInputBindings bindings) {
    WorkbookExecutionEngine workbookEngine = new WorkbookExecutionEngine();
    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(workbookEngine);
    return new StepExecution(
        workbookEngine,
        selectorResolver,
        new AssertionExecutor(workbookEngine, selectorResolver),
        WorkbookTempFileFactory.rooted(root.resolve("scratch"))::createTempFile,
        bindings);
  }

  private static ExecutionResults responseSupport() {
    return new ExecutionResults(XssfWorkbook::close, materialized -> materialized.close());
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
