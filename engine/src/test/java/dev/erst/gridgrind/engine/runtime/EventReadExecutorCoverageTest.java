package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Coverage for EVENT_READ execution mode paths in the default request executor. */
class EventReadExecutorCoverageTest {
  @Test
  void eventReadExecutionReturnsStructuredStepFailureForInspectionErrors() throws IOException {
    Path workbookPath = createWorkbookFile("gridgrind-event-step-failure-");

    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    ExecutionModeSpec.eventRead(),
                    null,
                    List.of(),
                    List.of(
                        inspect(
                            "missing-sheet",
                            new SheetSelector.ByName(new SheetName("Missing")),
                            new SheetIntrospectionQuery.GetSheetSummary())))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().type());
    assertEquals(
        java.util.Optional.of("Missing"),
        DefaultGridGrindRequestExecutorTestSupport.executeStepContext(failure)
            .sheetName()
            .map(dev.erst.gridgrind.core.vocabulary.WireString::value));
  }

  @Test
  void eventReadExecutionReturnsInspectionResultsWhenAllStepsSucceed() throws IOException {
    Path workbookPath = createWorkbookFile("gridgrind-event-success-");

    WorkbookResult.Success success =
        assertInstanceOf(
            WorkbookResult.Success.class,
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    ExecutionModeSpec.eventRead(),
                    null,
                    List.of(),
                    List.of(
                        inspect(
                            "summary",
                            new WorkbookSelector.Current(),
                            new WorkbookIntrospectionQuery.GetWorkbookSummary())))));

    assertEquals(1, success.inspections().size());
    assertEquals("summary", success.inspections().getFirst().stepId().value());
  }

  @Test
  void eventReadCloseFailureTurnsSuccessIntoExecuteRequestFailure() throws IOException {
    Path workbookPath = createWorkbookFile("gridgrind-event-close-success-failure-");

    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(
                    new DefaultGridGrindRequestExecutorDependencies(
                        new WorkbookExecutionEngine(),
                        ExcelWorkbook::close,
                        ignored -> {
                          throw new IOException("close failed");
                        },
                        dev.erst.gridgrind.excel.stream.ExcelStreamingWorkbookWriter
                            ::markRecalculateOnOpen)),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    ExecutionModeSpec.eventRead(),
                    null,
                    List.of(),
                    List.of(
                        inspect(
                            "summary",
                            new WorkbookSelector.Current(),
                            new WorkbookIntrospectionQuery.GetWorkbookSummary())))));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("EXECUTE_REQUEST", failure.problem().context().type());
  }

  @Test
  void eventReadCloseFailureAppendsCauseToExistingStepFailure() throws IOException {
    Path workbookPath = createWorkbookFile("gridgrind-event-close-step-failure-");

    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(
                    new DefaultGridGrindRequestExecutorDependencies(
                        new WorkbookExecutionEngine(),
                        ExcelWorkbook::close,
                        ignored -> {
                          throw new IOException("close failed");
                        },
                        dev.erst.gridgrind.excel.stream.ExcelStreamingWorkbookWriter
                            ::markRecalculateOnOpen)),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    ExecutionModeSpec.eventRead(),
                    null,
                    List.of(),
                    List.of(
                        inspect(
                            "missing-sheet",
                            new SheetSelector.ByName(new SheetName("Missing")),
                            new SheetIntrospectionQuery.GetSheetSummary())))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals(2, failure.problem().causes().size());
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().causes().get(1).code());
  }

  @Test
  void eventReadExecutionReturnsOpenWorkbookFailureWhenMaterializationFails() {
    Path missingWorkbook =
        ExecutionContextFixtureSupport.defaultWorkingDirectory()
            .resolve("gridgrind-missing-event-" + System.nanoTime() + ".xlsx");

    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.ExistingFile(missingWorkbook.toString()),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    ExecutionModeSpec.eventRead(),
                    null,
                    List.of(),
                    List.of(
                        inspect(
                            "summary",
                            new WorkbookSelector.Current(),
                            new WorkbookIntrospectionQuery.GetWorkbookSummary())))));

    assertEquals(GridGrindProblemCode.WORKBOOK_NOT_FOUND, failure.problem().code());
    assertEquals("OPEN_WORKBOOK", failure.problem().context().type());
  }
}
