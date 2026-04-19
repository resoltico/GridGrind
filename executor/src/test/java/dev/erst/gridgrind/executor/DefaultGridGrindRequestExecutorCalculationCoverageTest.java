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
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.CalculationExecutionStatus;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
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
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.nio.file.Files;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Focused Phase 6 coverage for calculation-policy paths in the default executor. */
@SuppressWarnings("PMD.AvoidAccessibilityAlteration")
class DefaultGridGrindRequestExecutorCalculationCoverageTest {
  @Test
  void executeRejectsMutationObservationOrderingAtValidationPhase() {
    WorkbookPlan request =
        new WorkbookPlan(
            GridGrindProtocolVersion.current(),
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            executionPolicy(calculateAll()),
            null,
            java.util.List.<WorkbookStep>of(
                new MutationStep(
                    "step-01-ensure-sheet",
                    new SheetSelector.ByName("Ops"),
                    new MutationAction.EnsureSheet()),
                new InspectionStep(
                    "summary",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookSummary()),
                new MutationStep(
                    "step-03-set-cell",
                    new CellSelector.ByAddress("Ops", "A1"),
                    new MutationAction.SetCell(
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
            new WorkbookCommandExecutor(),
            new WorkbookReadExecutor(),
            ExcelWorkbook::close,
            Files::createTempFile,
            ignored -> {},
            writer -> {
              throw new IllegalStateException("streaming calculation failed");
            });

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
                        mutate(new SheetSelector.ByName("Ops"), new MutationAction.EnsureSheet())),
                    java.util.List.of())));

    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.problem().code());
    assertEquals("CALCULATION_EXECUTION", failure.problem().context().stage());
    assertEquals(CalculationExecutionStatus.FAILED, failure.calculation().execution().status());
    assertTrue(
        failure.calculation().execution().message().contains("streaming calculation failed"));
  }

  @Test
  void calculationHelpersMapPreflightExecutionAndContextDetails() throws Exception {
    Method executeCalculationPolicy =
        DefaultGridGrindRequestExecutor.class.getDeclaredMethod(
            "executeCalculationPolicy",
            ExcelWorkbook.class,
            WorkbookPlan.class,
            ExecutionJournalRecorder.class);
    Method calculationProblemFor =
        DefaultGridGrindRequestExecutor.class.getDeclaredMethod(
            "calculationProblemFor",
            WorkbookPlan.class,
            CalculationPolicyExecutor.FailureDetail.class);
    Method calculationContextFor =
        DefaultGridGrindRequestExecutor.class.getDeclaredMethod(
            "calculationContextFor",
            WorkbookPlan.class,
            String.class,
            CalculationPolicyExecutor.FailureDetail.class);
    Method stageFor =
        DefaultGridGrindRequestExecutor.class.getDeclaredMethod(
            "stageFor", CalculationPolicyExecutor.Phase.class);
    executeCalculationPolicy.setAccessible(true);
    calculationProblemFor.setAccessible(true);
    calculationContextFor.setAccessible(true);
    stageFor.setAccessible(true);

    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
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

      Object preflightOutcome =
          executeCalculationPolicy.invoke(
              executor,
              workbook,
              preflightRequest,
              ExecutionJournalRecorder.start(preflightRequest, ExecutionJournalSink.NOOP));

      CalculationReport preflightReport = reportFrom(preflightOutcome);
      GridGrindResponse.Problem preflightFailure = failureFrom(preflightOutcome);
      assertEquals(CalculationExecutionStatus.FAILED, preflightReport.execution().status());
      assertEquals(
          GridGrindProblemCode.UNREGISTERED_USER_DEFINED_FUNCTION, preflightFailure.code());
      assertEquals("CALCULATION_PREFLIGHT", preflightFailure.context().stage());
      assertEquals("Ops", preflightFailure.context().sheetName());
      assertEquals("A1", preflightFailure.context().address());
      assertEquals("TEXTAFTER(\"a,b\",\",\")", preflightFailure.context().formula());
    }

    WorkbookPlan clearCachesRequest =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            executionPolicy(clearFormulaCaches()),
            null,
            java.util.List.of());
    try (ExcelWorkbook workbook = instantiateWorkbook(new IteratorFailingWorkbook())) {
      Object executionOutcome =
          executeCalculationPolicy.invoke(
              executor,
              workbook,
              clearCachesRequest,
              ExecutionJournalRecorder.start(clearCachesRequest, ExecutionJournalSink.NOOP));

      CalculationReport executionReport = reportFrom(executionOutcome);
      GridGrindResponse.Problem executionFailure = failureFrom(executionOutcome);
      assertEquals(CalculationExecutionStatus.FAILED, executionReport.execution().status());
      assertEquals(GridGrindProblemCode.INTERNAL_ERROR, executionFailure.code());
      assertEquals("CALCULATION_EXECUTION", executionFailure.context().stage());
      assertNull(executionFailure.context().sheetName());
      assertNull(executionFailure.context().address());
      assertNull(executionFailure.context().formula());
    }

    WorkbookPlan markOnlyRequest =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            executionPolicy(markRecalculateOnOpen()),
            null,
            java.util.List.of());
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      Object markOnlyOutcome =
          executeCalculationPolicy.invoke(
              executor,
              workbook,
              markOnlyRequest,
              ExecutionJournalRecorder.start(markOnlyRequest, ExecutionJournalSink.NOOP));

      CalculationReport markOnlyReport = reportFrom(markOnlyOutcome);
      assertEquals(CalculationExecutionStatus.SUCCEEDED, markOnlyReport.execution().status());
      assertTrue(markOnlyReport.execution().markRecalculateOnOpenApplied());
      assertNull(failureFrom(markOnlyOutcome));
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
    GridGrindResponse.Problem mappedCodeFailure =
        (GridGrindResponse.Problem)
            calculationProblemFor.invoke(executor, preflightRequest, codeFailure);
    assertEquals(GridGrindProblemCode.INVALID_FORMULA, mappedCodeFailure.code());
    assertEquals("CALCULATION_PREFLIGHT", mappedCodeFailure.context().stage());

    GridGrindResponse.ProblemContext.ExecuteCalculation nullFailureContext =
        (GridGrindResponse.ProblemContext.ExecuteCalculation)
            calculationContextFor.invoke(
                executor, preflightRequest, "CALCULATION_EXECUTION", (Object) null);
    assertEquals("CALCULATION_EXECUTION", nullFailureContext.stage());
    assertNull(nullFailureContext.sheetName());
    assertNull(nullFailureContext.address());
    assertNull(nullFailureContext.formula());

    assertEquals(
        "CALCULATION_PREFLIGHT", stageFor.invoke(null, CalculationPolicyExecutor.Phase.PREFLIGHT));
    assertEquals(
        "CALCULATION_EXECUTION", stageFor.invoke(null, CalculationPolicyExecutor.Phase.EXECUTION));
  }

  private static GridGrindResponse.Failure failure(GridGrindResponse response) {
    return assertInstanceOf(GridGrindResponse.Failure.class, response);
  }

  private static CalculationReport reportFrom(Object outcome) {
    try {
      Method report = outcome.getClass().getDeclaredMethod("report");
      report.setAccessible(true);
      return (CalculationReport) report.invoke(outcome);
    } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException exception) {
      throw new LinkageError(exception.getMessage(), exception);
    }
  }

  private static GridGrindResponse.Problem failureFrom(Object outcome) {
    try {
      Method failure = outcome.getClass().getDeclaredMethod("failure");
      failure.setAccessible(true);
      return (GridGrindResponse.Problem) failure.invoke(outcome);
    } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException exception) {
      throw new LinkageError(exception.getMessage(), exception);
    }
  }

  private static ExcelWorkbook instantiateWorkbook(XSSFWorkbook workbook) {
    try {
      Class<?> runtimeType = Class.forName("dev.erst.gridgrind.excel.ExcelFormulaRuntime");
      Method runtimeFactory = runtimeType.getDeclaredMethod("poi", FormulaEvaluator.class);
      runtimeFactory.setAccessible(true);
      Object runtime =
          runtimeFactory.invoke(null, workbook.getCreationHelper().createFormulaEvaluator());
      Constructor<ExcelWorkbook> constructor =
          ExcelWorkbook.class.getDeclaredConstructor(XSSFWorkbook.class, runtimeType);
      constructor.setAccessible(true);
      return constructor.newInstance(workbook, runtime);
    } catch (ClassNotFoundException
        | NoSuchMethodException
        | IllegalAccessException
        | InstantiationException
        | InvocationTargetException exception) {
      throw new LinkageError(exception.getMessage(), exception);
    }
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
