package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.*;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindResponsePersistence;
import dev.erst.gridgrind.contract.dto.GridGrindWorkbookSurfaceReports;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.excel.ExcelBorderSideSnapshot;
import dev.erst.gridgrind.excel.ExcelBorderSnapshot;
import dev.erst.gridgrind.excel.ExcelCellAlignmentSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFillSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFontSnapshot;
import dev.erst.gridgrind.excel.ExcelCellProtectionSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelPrintMarginsSnapshot;
import dev.erst.gridgrind.excel.ExcelPrintSetupSnapshot;
import dev.erst.gridgrind.excel.ExcelSheetDefaults;
import dev.erst.gridgrind.excel.ExcelSheetDisplay;
import dev.erst.gridgrind.excel.ExcelSheetOutlineSummary;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.foundation.ExcelBorderStyle;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.math.BigDecimal;
import java.util.List;

/** Shared helpers for DefaultGridGrindRequestExecutor integration tests. */
class DefaultGridGrindRequestExecutorTestSupport {
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
      InspectionStep... inspections) {
    return ExecutorTestPlanSupport.request(source, persistence, mutations, inspections);
  }

  static GridGrindResponse.Success success(GridGrindResponse response) {
    return cast(GridGrindResponse.Success.class, response);
  }

  static GridGrindResponse.Failure failure(GridGrindResponse response) {
    return cast(GridGrindResponse.Failure.class, response);
  }

  static ProblemContext.ReadRequest readRequestContext(GridGrindProblemDetail.Problem problem) {
    return cast(ProblemContext.ReadRequest.class, problem.context());
  }

  static ProblemContext.ReadRequest readRequestContext(GridGrindResponse.Failure failure) {
    return readRequestContext(failure.problem());
  }

  static ProblemContext.OpenWorkbook openWorkbookContext(GridGrindProblemDetail.Problem problem) {
    return cast(ProblemContext.OpenWorkbook.class, problem.context());
  }

  static ProblemContext.OpenWorkbook openWorkbookContext(GridGrindResponse.Failure failure) {
    return openWorkbookContext(failure.problem());
  }

  static ProblemContext.PersistWorkbook persistWorkbookContext(
      GridGrindProblemDetail.Problem problem) {
    return cast(ProblemContext.PersistWorkbook.class, problem.context());
  }

  static ProblemContext.PersistWorkbook persistWorkbookContext(GridGrindResponse.Failure failure) {
    return persistWorkbookContext(failure.problem());
  }

  static ProblemContext.ExecuteRequest executeRequestContext(
      GridGrindProblemDetail.Problem problem) {
    return cast(ProblemContext.ExecuteRequest.class, problem.context());
  }

  static ProblemContext.ExecuteRequest executeRequestContext(GridGrindResponse.Failure failure) {
    return executeRequestContext(failure.problem());
  }

  static ProblemContext.ExecuteStep executeStepContext(GridGrindProblemDetail.Problem problem) {
    return cast(ProblemContext.ExecuteStep.class, problem.context());
  }

  static ProblemContext.ExecuteStep executeStepContext(GridGrindResponse.Failure failure) {
    return executeStepContext(failure.problem());
  }

  static ProblemContext.ExecuteCalculation.Preflight calculationPreflightContext(
      GridGrindProblemDetail.Problem problem) {
    return cast(ProblemContext.ExecuteCalculation.Preflight.class, problem.context());
  }

  static ProblemContext.ExecuteCalculation.Preflight calculationPreflightContext(
      GridGrindResponse.Failure failure) {
    return calculationPreflightContext(failure.problem());
  }

  static ProblemContext.ExecuteCalculation.Execution calculationExecutionContext(
      GridGrindProblemDetail.Problem problem) {
    return cast(ProblemContext.ExecuteCalculation.Execution.class, problem.context());
  }

  static ProblemContext.ExecuteCalculation.Execution calculationExecutionContext(
      GridGrindResponse.Failure failure) {
    return calculationExecutionContext(failure.problem());
  }

  static String savedPath(GridGrindResponse.Success success) {
    return switch (success.persistence()) {
      case GridGrindResponsePersistence.PersistenceOutcome.SavedAs savedAs ->
          savedAs.executionPath();
      case GridGrindResponsePersistence.PersistenceOutcome.Overwritten overwritten ->
          overwritten.executionPath();
      case GridGrindResponsePersistence.PersistenceOutcome.NotSaved _ ->
          throw new AssertionError("expected persisted workbook");
    };
  }

  static List<String> stepIds(GridGrindResponse.Success success) {
    return inspectionIds(success);
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

  static <T> T cast(Class<T> type, Object value) {
    return type.cast(assertInstanceOf(type, value));
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

  static <T extends InspectionResult> T read(
      GridGrindResponse.Success success, String stepId, Class<T> type) {
    return inspection(success, stepId, type);
  }

  GridGrindWorkbookSurfaceReports.CellStyleReport toResponseStyleReport(
      dev.erst.gridgrind.excel.ExcelCellStyleSnapshot style) {
    return InspectionResultCellReportSupport.toCellStyleReport(style);
  }

  static ExcelCellStyleSnapshot defaultStyle() {
    return new ExcelCellStyleSnapshot(
        "",
        new ExcelCellAlignmentSnapshot(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new ExcelCellFontSnapshot(
            false,
            false,
            "Aptos",
            ExcelFontHeight.fromPoints(new BigDecimal("11")),
            null,
            false,
            false),
        ExcelCellFillSnapshot.pattern(dev.erst.gridgrind.excel.foundation.ExcelFillPattern.NONE),
        new ExcelBorderSnapshot(
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null),
            new ExcelBorderSideSnapshot(ExcelBorderStyle.NONE, null)),
        new ExcelCellProtectionSnapshot(true, false));
  }

  static ExcelPrintSetupSnapshot defaultPrintSetupSnapshot() {
    return new ExcelPrintSetupSnapshot(
        new ExcelPrintMarginsSnapshot(0.0d, 0.0d, 0.0d, 0.0d, 0.0d, 0.0d),
        false,
        false,
        false,
        0,
        false,
        false,
        0,
        false,
        0,
        List.of(),
        List.of());
  }

  static dev.erst.gridgrind.excel.ExcelSheetPresentationSnapshot
      defaultSheetPresentationSnapshot() {
    return new dev.erst.gridgrind.excel.ExcelSheetPresentationSnapshot(
        ExcelSheetDisplay.defaults(),
        null,
        ExcelSheetOutlineSummary.defaults(),
        ExcelSheetDefaults.defaults(),
        List.of());
  }

  static SheetProtectionSettings protectionSettings() {
    return new SheetProtectionSettings(
        false, true, false, true, false, true, false, true, false, true, false, true, false, true,
        false);
  }

  static CellColorReport rgb(String rgb) {
    return CellColorReport.rgb(rgb);
  }

  static ExcelSheetProtectionSettings excelProtectionSettings() {
    return new ExcelSheetProtectionSettings(
        false, true, false, true, false, true, false, true, false, true, false, true, false, true,
        false);
  }
}
