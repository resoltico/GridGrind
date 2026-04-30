package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.calculateAll;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.clearFormulaCaches;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.executionPolicy;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.inspect;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.markRecalculateOnOpen;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutate;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.request;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.CalculationExecutionStatus;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelStreamingWorkbookWriter;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import java.nio.file.Files;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Focused Phase 6 coverage for calculation-policy paths in the default executor. */
class DefaultGridGrindRequestExecutorCalculationCoverageTest {
  @Test
  void executeRejectsMutationObservationOrderingAtValidationPhase() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            executionPolicy(calculateAll()),
            FormulaEnvironmentInput.empty(),
            java.util.List.<WorkbookStep>of(
                new MutationStep(
                    "step-01-ensure-sheet",
                    new SheetSelector.ByName("Ops"),
                    new WorkbookMutationAction.EnsureSheet()),
                new InspectionStep(
                    "summary",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookSummary()),
                new MutationStep(
                    "step-03-set-cell",
                    new CellSelector.ByAddress("Ops", "A1"),
                    new CellMutationAction.SetCell(
                        new dev.erst.gridgrind.contract.dto.CellInput.Numeric(1.0)))));
    GridGrindResponse.Failure failure =
        failure(new DefaultGridGrindRequestExecutor().execute(request));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("VALIDATE_REQUEST", failure.problem().context().stage());
    assertTrue(failure.problem().message().contains("mutation-to-observation boundary"));
  }

  @Test
  void directEventReadEligibilityRejectsCalculationPoliciesThatNeedExecution() {
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.ExistingFile("/tmp/calculation-event-read.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None(),
            executionPolicy(
                new ExecutionModeInput(
                    ExecutionModeInput.ReadMode.EVENT_READ, ExecutionModeInput.WriteMode.FULL_XSSF),
                calculateAll()),
            null,
            java.util.List.of(),
            java.util.List.of(
                inspect(
                    "summary",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookSummary())));

    assertFalse(
        DefaultGridGrindRequestExecutor.directEventReadEligible(
            request, DefaultGridGrindRequestExecutor.executionModes(request)));
  }

  @Test
  void streamingCalculationRuntimeFailureReturnsCalculationFailureBeforePersistence() {
    DefaultGridGrindRequestExecutor executor =
        new DefaultGridGrindRequestExecutor(
            new DefaultGridGrindRequestExecutorDependencies(
                new WorkbookCommandExecutor(),
                new WorkbookReadExecutor(),
                ExcelWorkbook::close,
                Files::createTempFile,
                ignored -> {},
                writer -> {
                  throw new IllegalStateException("streaming calculation failed");
                }));

    GridGrindResponse.Failure failure =
        failure(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    executionPolicy(
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.FULL_XSSF,
                            ExecutionModeInput.WriteMode.STREAMING_WRITE),
                        markRecalculateOnOpen()),
                    null,
                    java.util.List.of(
                        mutate(
                            new SheetSelector.ByName("Ops"),
                            new WorkbookMutationAction.EnsureSheet())),
                    java.util.List.of())));

    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.problem().code());
    assertEquals("CALCULATION_EXECUTION", failure.problem().context().stage());
    assertEquals(CalculationExecutionStatus.FAILED, failure.calculation().execution().status());
    assertTrue(
        failure
            .calculation()
            .execution()
            .message()
            .orElseThrow()
            .contains("streaming calculation failed"));
  }

  @Test
  void calculationHelpersMapPreflightExecutionAndContextDetails() throws Exception {
    ExecutionCalculationSupport calculationSupport =
        new ExecutionCalculationSupport(ExcelStreamingWorkbookWriter::markRecalculateOnOpen);
    WorkbookPlan preflightRequest =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            executionPolicy(calculateAll()),
            null,
            java.util.List.of());
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.sheet("Ops").setCell("A1", ExcelCellValue.formula("TEXTAFTER(\"a,b\",\",\")"));

      ExecutionCalculationSupport.CalculationExecutionOutcome preflightOutcome =
          calculationSupport.executeCalculationPolicy(
              workbook,
              preflightRequest,
              ExecutionJournalRecorder.start(preflightRequest, ExecutionJournalSink.NOOP));

      CalculationReport preflightReport = preflightOutcome.report();
      GridGrindProblemDetail.Problem preflightFailure = preflightOutcome.failure().orElseThrow();
      assertEquals(CalculationExecutionStatus.FAILED, preflightReport.execution().status());
      assertEquals(
          GridGrindProblemCode.UNREGISTERED_USER_DEFINED_FUNCTION, preflightFailure.code());
      assertEquals("CALCULATION_PREFLIGHT", preflightFailure.context().stage());
      assertEquals(
          java.util.Optional.of("Ops"),
          DefaultGridGrindRequestExecutorTestSupport.calculationPreflightContext(preflightFailure)
              .sheetName());
      assertEquals(
          java.util.Optional.of("A1"),
          DefaultGridGrindRequestExecutorTestSupport.calculationPreflightContext(preflightFailure)
              .address());
      assertEquals(
          java.util.Optional.of("TEXTAFTER(\"a,b\",\",\")"),
          DefaultGridGrindRequestExecutorTestSupport.calculationPreflightContext(preflightFailure)
              .formula());
    }

    WorkbookPlan clearCachesRequest =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            executionPolicy(clearFormulaCaches()),
            null,
            java.util.List.of());
    try (ExcelWorkbook workbook = instantiateWorkbook(new IteratorFailingWorkbook())) {
      ExecutionCalculationSupport.CalculationExecutionOutcome executionOutcome =
          calculationSupport.executeCalculationPolicy(
              workbook,
              clearCachesRequest,
              ExecutionJournalRecorder.start(clearCachesRequest, ExecutionJournalSink.NOOP));

      CalculationReport executionReport = executionOutcome.report();
      GridGrindProblemDetail.Problem executionFailure = executionOutcome.failure().orElseThrow();
      assertEquals(CalculationExecutionStatus.FAILED, executionReport.execution().status());
      assertEquals(GridGrindProblemCode.INTERNAL_ERROR, executionFailure.code());
      assertEquals("CALCULATION_EXECUTION", executionFailure.context().stage());
      assertEquals(
          java.util.Optional.empty(),
          DefaultGridGrindRequestExecutorTestSupport.calculationExecutionContext(executionFailure)
              .sheetName());
      assertEquals(
          java.util.Optional.empty(),
          DefaultGridGrindRequestExecutorTestSupport.calculationExecutionContext(executionFailure)
              .address());
      assertEquals(
          java.util.Optional.empty(),
          DefaultGridGrindRequestExecutorTestSupport.calculationExecutionContext(executionFailure)
              .formula());
    }

    WorkbookPlan markOnlyRequest =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            executionPolicy(markRecalculateOnOpen()),
            null,
            java.util.List.of());
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExecutionCalculationSupport.CalculationExecutionOutcome markOnlyOutcome =
          calculationSupport.executeCalculationPolicy(
              workbook,
              markOnlyRequest,
              ExecutionJournalRecorder.start(markOnlyRequest, ExecutionJournalSink.NOOP));

      CalculationReport markOnlyReport = markOnlyOutcome.report();
      assertEquals(CalculationExecutionStatus.SUCCEEDED, markOnlyReport.execution().status());
      assertTrue(markOnlyReport.execution().markRecalculateOnOpenApplied());
      assertTrue(markOnlyOutcome.failure().isEmpty());
    }

    CalculationPolicyExecutor.FailureDetail codeFailure =
        new CalculationPolicyExecutor.FailureDetail(
            GridGrindProblemCode.INVALID_FORMULA,
            CalculationPolicyExecutor.Phase.PREFLIGHT,
            "Ops",
            "B2",
            "SUM(",
            "invalid formula",
            null);
    GridGrindProblemDetail.Problem mappedCodeFailure =
        calculationSupport.calculationProblemFor(preflightRequest, codeFailure);
    assertEquals(GridGrindProblemCode.INVALID_FORMULA, mappedCodeFailure.code());
    assertEquals("CALCULATION_PREFLIGHT", mappedCodeFailure.context().stage());

    dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteCalculation nullFailureContext =
        calculationSupport.calculationContextFor(
            preflightRequest, CalculationPolicyExecutor.Phase.EXECUTION, null);
    assertEquals("CALCULATION_EXECUTION", nullFailureContext.stage());
    assertEquals(java.util.Optional.empty(), nullFailureContext.sheetName());
    assertEquals(java.util.Optional.empty(), nullFailureContext.address());
    assertEquals(java.util.Optional.empty(), nullFailureContext.formula());

    dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteCalculation preflightNullFailureContext =
        calculationSupport.calculationContextFor(
            preflightRequest, CalculationPolicyExecutor.Phase.PREFLIGHT, null);
    assertEquals("CALCULATION_PREFLIGHT", preflightNullFailureContext.stage());
    assertEquals(java.util.Optional.empty(), preflightNullFailureContext.sheetName());
    assertEquals(java.util.Optional.empty(), preflightNullFailureContext.address());
    assertEquals(java.util.Optional.empty(), preflightNullFailureContext.formula());

    CalculationPolicyExecutor.FailureDetail missingFormulaField =
        new CalculationPolicyExecutor.FailureDetail(
            GridGrindProblemCode.INVALID_FORMULA,
            CalculationPolicyExecutor.Phase.EXECUTION,
            "Budget",
            "B4",
            null,
            "invalid formula",
            null);
    dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteCalculation missingFormulaContext =
        calculationSupport.calculationContextFor(
            preflightRequest, CalculationPolicyExecutor.Phase.EXECUTION, missingFormulaField);
    assertEquals(java.util.Optional.empty(), missingFormulaContext.sheetName());
    assertEquals(java.util.Optional.empty(), missingFormulaContext.address());

    CalculationPolicyExecutor.FailureDetail missingAddressField =
        new CalculationPolicyExecutor.FailureDetail(
            GridGrindProblemCode.INVALID_FORMULA,
            CalculationPolicyExecutor.Phase.EXECUTION,
            "Budget",
            null,
            "SUM(",
            "invalid formula",
            null);
    dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteCalculation missingAddressContext =
        calculationSupport.calculationContextFor(
            preflightRequest, CalculationPolicyExecutor.Phase.EXECUTION, missingAddressField);
    assertEquals(java.util.Optional.empty(), missingAddressContext.sheetName());
  }

  private static GridGrindResponse.Failure failure(GridGrindResponse response) {
    return assertInstanceOf(GridGrindResponse.Failure.class, response);
  }

  private static ExcelWorkbook instantiateWorkbook(XSSFWorkbook workbook) {
    return ExcelWorkbook.wrap(workbook);
  }

  /**
   * Workbook whose sheet iteration path fails so executor calculation error mapping is exercised.
   */
  private static final class IteratorFailingWorkbook extends XSSFWorkbook {
    @Override
    public Iterator<Sheet> iterator() {
      throw new IllegalStateException("clear caches failed");
    }
  }
}
