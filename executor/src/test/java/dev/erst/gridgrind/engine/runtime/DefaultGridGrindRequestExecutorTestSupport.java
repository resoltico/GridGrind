package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.WorkbookTempFileFactory;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;

/** Shared helpers for DefaultGridGrindRequestExecutor integration tests. */
class DefaultGridGrindRequestExecutorTestSupport
    extends DefaultGridGrindRequestExecutorReadSupport {
  protected DefaultGridGrindRequestExecutorTestSupport() {}

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<ExecutorTestPlanSupport.PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return ExecutorTestPlanSupport.request(
        source, persistence, execution, formulaEnvironment, mutations, inspections);
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<ExecutorTestPlanSupport.PendingMutation> mutations,
      List<ExecutorTestPlanSupport.PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    return ExecutorTestPlanSupport.request(
        source, persistence, execution, formulaEnvironment, mutations, assertions, inspections);
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionPolicyInput execution,
      FormulaEnvironmentInput formulaEnvironment,
      List<ExecutorTestPlanSupport.PendingMutation> mutations,
      InspectionStep... inspections) {
    return ExecutorTestPlanSupport.request(
        source, persistence, execution, formulaEnvironment, mutations, inspections);
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      List<ExecutorTestPlanSupport.PendingMutation> mutations,
      List<ExecutorTestPlanSupport.PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    return ExecutorTestPlanSupport.request(source, persistence, mutations, assertions, inspections);
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      List<ExecutorTestPlanSupport.PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return ExecutorTestPlanSupport.request(source, persistence, mutations, inspections);
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      List<ExecutorTestPlanSupport.PendingMutation> mutations,
      InspectionStep... inspections) {
    return ExecutorTestPlanSupport.request(source, persistence, mutations, inspections);
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      FormulaEnvironmentInput formulaEnvironment,
      List<ExecutorTestPlanSupport.PendingMutation> mutations,
      List<ExecutorTestPlanSupport.PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    return ExecutorTestPlanSupport.request(
        source, persistence, formulaEnvironment, mutations, assertions, inspections);
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      FormulaEnvironmentInput formulaEnvironment,
      List<ExecutorTestPlanSupport.PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return ExecutorTestPlanSupport.request(
        source, persistence, formulaEnvironment, mutations, inspections);
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionModeInput executionMode,
      FormulaEnvironmentInput formulaEnvironment,
      List<ExecutorTestPlanSupport.PendingMutation> mutations,
      List<ExecutorTestPlanSupport.PendingAssertion> assertions,
      List<InspectionStep> inspections) {
    return ExecutorTestPlanSupport.request(
        source, persistence, executionMode, formulaEnvironment, mutations, assertions, inspections);
  }

  static WorkbookPlan request(
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence,
      ExecutionModeInput executionMode,
      FormulaEnvironmentInput formulaEnvironment,
      List<ExecutorTestPlanSupport.PendingMutation> mutations,
      List<InspectionStep> inspections) {
    return ExecutorTestPlanSupport.request(
        source, persistence, executionMode, formulaEnvironment, mutations, inspections);
  }

  static WorkbookResult.Success success(WorkbookResult response) {
    return cast(WorkbookResult.Success.class, response);
  }

  static WorkbookResult.Failure failure(WorkbookResult response) {
    return cast(WorkbookResult.Failure.class, response);
  }

  static ProblemContext.ReadRequest readRequestContext(GridGrindProblemDetail.Problem problem) {
    return cast(ProblemContext.ReadRequest.class, problem.context());
  }

  static ProblemContext.ReadRequest readRequestContext(WorkbookResult.Failure failure) {
    return readRequestContext(failure.problem());
  }

  static ProblemContext.OpenWorkbook openWorkbookContext(GridGrindProblemDetail.Problem problem) {
    return cast(ProblemContext.OpenWorkbook.class, problem.context());
  }

  static ProblemContext.OpenWorkbook openWorkbookContext(WorkbookResult.Failure failure) {
    return openWorkbookContext(failure.problem());
  }

  static ProblemContext.PersistWorkbook persistWorkbookContext(
      GridGrindProblemDetail.Problem problem) {
    return cast(ProblemContext.PersistWorkbook.class, problem.context());
  }

  static ProblemContext.PersistWorkbook persistWorkbookContext(WorkbookResult.Failure failure) {
    return persistWorkbookContext(failure.problem());
  }

  static ProblemContext.ExecuteRequest executeRequestContext(
      GridGrindProblemDetail.Problem problem) {
    return cast(ProblemContext.ExecuteRequest.class, problem.context());
  }

  static ProblemContext.ExecuteRequest executeRequestContext(WorkbookResult.Failure failure) {
    return executeRequestContext(failure.problem());
  }

  static ProblemContext.ExecuteStep executeStepContext(GridGrindProblemDetail.Problem problem) {
    return cast(ProblemContext.ExecuteStep.class, problem.context());
  }

  static ProblemContext.ExecuteStep executeStepContext(WorkbookResult.Failure failure) {
    return executeStepContext(failure.problem());
  }

  static ProblemContext.ExecuteCalculation.Preflight calculationPreflightContext(
      GridGrindProblemDetail.Problem problem) {
    return cast(ProblemContext.ExecuteCalculation.Preflight.class, problem.context());
  }

  static ProblemContext.ExecuteCalculation.Preflight calculationPreflightContext(
      WorkbookResult.Failure failure) {
    return calculationPreflightContext(failure.problem());
  }

  static ProblemContext.ExecuteCalculation.Execution calculationExecutionContext(
      GridGrindProblemDetail.Problem problem) {
    return cast(ProblemContext.ExecuteCalculation.Execution.class, problem.context());
  }

  static ProblemContext.ExecuteCalculation.Execution calculationExecutionContext(
      WorkbookResult.Failure failure) {
    return calculationExecutionContext(failure.problem());
  }

  static String savedPath(WorkbookResult.Success success) {
    return writtenExecutionPath(success.persistence());
  }

  static String writtenExecutionPath(WorkbookResultPersistence.PersistenceOutcome persistence) {
    return switch (persistence) {
      case WorkbookResultPersistence.PersistenceOutcome.SavedAs savedAs ->
          writtenExecutionPath(savedAs);
      case WorkbookResultPersistence.PersistenceOutcome.Overwritten overwritten ->
          writtenExecutionPath(overwritten);
      case WorkbookResultPersistence.PersistenceOutcome.NotSaved _ ->
          throw new AssertionError("expected persisted workbook");
    };
  }

  static String writtenExecutionPath(WorkbookResultPersistence.PersistenceOutcome.SavedAs savedAs) {
    return writtenExecutionPath(savedAs.write());
  }

  static String writtenExecutionPath(
      WorkbookResultPersistence.PersistenceOutcome.Overwritten overwritten) {
    return writtenExecutionPath(overwritten.write());
  }

  static String writtenExecutionPath(WorkbookResultPersistence.WriteResult write) {
    return switch (write) {
      case WorkbookResultPersistence.WriteResult.Written written -> written.executionPath();
      case WorkbookResultPersistence.WriteResult.NotWritten _ ->
          throw new AssertionError("expected written workbook");
    };
  }

  static List<String> stepIds(WorkbookResult.Success success) {
    return inspectionIds(success);
  }

  static ExecutionInputBindings bindings(Path workingDirectory) {
    return ExecutionInputBindingsFixtureSupport.bindings(workingDirectory);
  }

  static ExecutionInputBindings defaultBindings() {
    return ExecutionContextFixtureSupport.defaultBindings();
  }

  static WorkbookTempFileFactory tempFileFactory(Path workingDirectory) {
    return ExecutionContextFixtureSupport.tempFileFactory(workingDirectory);
  }

  static WorkbookTempFileFactory tempFileFactoryFor(Path anchoredPath) {
    return ExecutionContextFixtureSupport.tempFileFactoryFor(anchoredPath);
  }

  static ExecutionWorkbookSupport workbookSupport(Path workingDirectory) {
    return ExecutionContextFixtureSupport.workbookSupport(workingDirectory);
  }

  static ExecutionJournalRecorder startJournal(WorkbookPlan request) {
    return ExecutionContextFixtureSupport.startJournal(request);
  }

  static ExecutionJournalRecorder startJournal(
      WorkbookPlan request, ExecutionProgressSink sink, Path workingDirectory) {
    return ExecutionContextFixtureSupport.startJournal(request, sink, workingDirectory);
  }

  static WorkbookResult execute(DefaultGridGrindRequestExecutor executor, WorkbookPlan request) {
    return ExecutionContextFixtureSupport.execute(executor, request);
  }

  static WorkbookResult execute(
      DefaultGridGrindRequestExecutor executor, WorkbookPlan request, Path workingDirectory) {
    return ExecutionContextFixtureSupport.execute(executor, request, workingDirectory);
  }

  static void saveWorkbook(ExcelWorkbook workbook, Path workbookPath) throws IOException {
    ExecutionContextFixtureSupport.saveWorkbook(workbook, workbookPath);
  }

  static WorkbookCommand command(ExecutorTestPlanSupport.PendingMutation mutation) {
    return WorkbookCommandConverter.toCommand(mutation.target(), mutation.action());
  }

  static WorkbookReadCommand readCommand(InspectionStep step) {
    return InspectionCommandConverter.toReadCommand(step);
  }

  static String readType(InspectionStep step) {
    return step.query().queryType();
  }

  static String sheetNameFor(WorkbookStep step) {
    return ExecutionDiagnosticFields.sheetNameFor(step).orElse(null);
  }

  static String sheetNameFor(WorkbookStep step, Exception exception) {
    return ExecutionDiagnosticFields.sheetNameFor(step, exception).orElse(null);
  }

  static String sheetNameFor(
      ExecutorTestPlanSupport.PendingMutation mutation, Exception exception) {
    return sheetNameFor(materializeMutation(mutation, 0), exception);
  }

  static String addressFor(WorkbookStep step, Exception exception) {
    return ExecutionDiagnosticFields.addressFor(step, exception).orElse(null);
  }

  static String addressFor(ExecutorTestPlanSupport.PendingMutation mutation, Exception exception) {
    return addressFor(materializeMutation(mutation, 0), exception);
  }

  static String rangeFor(WorkbookStep step, Exception exception) {
    return ExecutionDiagnosticFields.rangeFor(step, exception).orElse(null);
  }

  static String rangeFor(ExecutorTestPlanSupport.PendingMutation mutation, Exception exception) {
    return rangeFor(materializeMutation(mutation, 0), exception);
  }

  static String formulaFor(WorkbookStep step, Exception exception) {
    return ExecutionDiagnosticFields.formulaFor(step, exception).orElse(null);
  }

  static String formulaFor(ExecutorTestPlanSupport.PendingMutation mutation, Exception exception) {
    return formulaFor(materializeMutation(mutation, 0), exception);
  }

  static String namedRangeNameFor(
      ExecutorTestPlanSupport.PendingMutation mutation, Exception exception) {
    return namedRangeNameFor(materializeMutation(mutation, 0), exception);
  }

  static String namedRangeNameFor(WorkbookStep step, Exception exception) {
    return ExecutionDiagnosticFields.namedRangeNameFor(step, exception).orElse(null);
  }

  static void assertReadContext(
      InspectionStep step,
      String expectedSheetName,
      String expectedRuntimeAddress,
      String expectedNamedRangeName,
      RuntimeException runtimeException) {
    assertEquals(expectedSheetName, sheetNameFor(step));
    assertEquals(expectedRuntimeAddress, addressFor(step, runtimeException));
    assertEquals(expectedNamedRangeName, namedRangeNameFor(step, runtimeException));
  }

  static void assertWriteContext(
      ExecutorTestPlanSupport.PendingMutation mutation,
      Exception exception,
      String expectedSheetName,
      String expectedAddress,
      String expectedRange,
      String expectedNamedRangeName) {
    MutationStep step = materializeMutation(mutation, 0);
    assertNull(formulaFor(step, exception));
    assertEquals(expectedSheetName, sheetNameFor(step, exception));
    assertEquals(expectedAddress, addressFor(step, exception));
    assertEquals(expectedRange, rangeFor(step, exception));
    assertEquals(expectedNamedRangeName, namedRangeNameFor(step, exception));
  }

  CellStyleReport toResponseStyleReport(dev.erst.gridgrind.excel.ExcelCellStyleSnapshot style) {
    return InspectionResultCellStyleReportSupport.toCellStyleReport(style);
  }
}
