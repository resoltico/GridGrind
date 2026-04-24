package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.DifferentialStyleInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.CellNotFoundException;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelOoxmlPackageSecuritySnapshot;
import dev.erst.gridgrind.excel.ExcelOoxmlPackageSecuritySupport;
import dev.erst.gridgrind.excel.ExcelOoxmlPersistenceOptions;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.MissingExternalWorkbookException;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.UnregisteredUserDefinedFunctionException;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlEncryptionMode;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Phase-3 coverage lock-in for executor policy seams, helper extraction, and failure routing. */
class ExecutorPhase3CoverageTest {
  @Test
  void inspectionCommandConverterRejectsUnsupportedChartTargets() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                InspectionCommandConverter.toReadCommand(
                    "charts", new WorkbookSelector.Current(), new InspectionQuery.GetCharts()));

    assertEquals("Unsupported chart inspection target", failure.getMessage());
  }

  @Test
  void workbookCommandConverterRejectsIdentityMismatchesAndBroadNamedRangeDeletes() {
    IllegalArgumentException tableMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookCommandConverter.toCommand(
                    new TableSelector.ByNameOnSheet("OtherTable", "Budget"),
                    new MutationAction.SetTable(
                        new TableInput(
                            "BudgetTable", "Budget", "A1:B2", false, new TableStyleInput.None()))));
    assertEquals(
        "SET_TABLE target must match table.name and table.sheetName", tableMismatch.getMessage());

    IllegalArgumentException tableSheetMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookCommandConverter.toCommand(
                    new TableSelector.ByNameOnSheet("BudgetTable", "Archive"),
                    new MutationAction.SetTable(
                        new TableInput(
                            "BudgetTable", "Budget", "A1:B2", false, new TableStyleInput.None()))));
    assertEquals(
        "SET_TABLE target must match table.name and table.sheetName",
        tableSheetMismatch.getMessage());

    IllegalArgumentException pivotMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookCommandConverter.toCommand(
                    new PivotTableSelector.ByNameOnSheet("OtherPivot", "Budget"),
                    new MutationAction.SetPivotTable(
                        new PivotTableInput(
                            "SalesPivot",
                            "Budget",
                            new PivotTableInput.Source.NamedRange("BudgetSource"),
                            new PivotTableInput.Anchor("B3"),
                            List.of(),
                            List.of(),
                            List.of(),
                            List.of(
                                new PivotTableInput.DataField(
                                    "Amount",
                                    dev.erst.gridgrind.excel.foundation
                                        .ExcelPivotDataConsolidateFunction.SUM,
                                    null,
                                    null))))));
    assertEquals(
        "SET_PIVOT_TABLE target must match pivotTable.name and pivotTable.sheetName",
        pivotMismatch.getMessage());

    IllegalArgumentException pivotSheetMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookCommandConverter.toCommand(
                    new PivotTableSelector.ByNameOnSheet("SalesPivot", "Archive"),
                    new MutationAction.SetPivotTable(
                        new PivotTableInput(
                            "SalesPivot",
                            "Budget",
                            new PivotTableInput.Source.NamedRange("BudgetSource"),
                            new PivotTableInput.Anchor("B3"),
                            List.of(),
                            List.of(),
                            List.of(),
                            List.of(
                                new PivotTableInput.DataField(
                                    "Amount",
                                    dev.erst.gridgrind.excel.foundation
                                        .ExcelPivotDataConsolidateFunction.SUM,
                                    null,
                                    null))))));
    assertEquals(
        "SET_PIVOT_TABLE target must match pivotTable.name and pivotTable.sheetName",
        pivotSheetMismatch.getMessage());

    IllegalArgumentException namedRangeMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookCommandConverter.toCommand(
                    new NamedRangeSelector.SheetScope("BudgetTotal", "Archive"),
                    new MutationAction.SetNamedRange(
                        "BudgetTotal",
                        new NamedRangeScope.Sheet("Budget"),
                        new NamedRangeTarget("Budget", "B4"))));
    assertEquals(
        "SET_NAMED_RANGE target must match action name and scope", namedRangeMismatch.getMessage());

    IllegalArgumentException namedRangeNameMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookCommandConverter.toCommand(
                    new NamedRangeSelector.SheetScope("OtherTotal", "Budget"),
                    new MutationAction.SetNamedRange(
                        "BudgetTotal",
                        new NamedRangeScope.Sheet("Budget"),
                        new NamedRangeTarget("Budget", "B4"))));
    assertEquals(
        "SET_NAMED_RANGE target must match action name and scope",
        namedRangeNameMismatch.getMessage());

    IllegalArgumentException workbookScopeMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookCommandConverter.toCommand(
                    new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                    new MutationAction.SetNamedRange(
                        "BudgetTotal",
                        new NamedRangeScope.Sheet("Budget"),
                        new NamedRangeTarget("Budget", "B4"))));
    assertEquals(
        "SET_NAMED_RANGE target must match action name and scope",
        workbookScopeMismatch.getMessage());

    IllegalArgumentException broadDelete =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookCommandConverter.toCommand(
                    new NamedRangeSelector.ByName("BudgetTotal"),
                    new MutationAction.DeleteNamedRange()));
    assertEquals(
        "DELETE_NAMED_RANGE requires target type ScopedExact but got ByName",
        broadDelete.getMessage());
  }

  @Test
  void gridGrindProblemsAndWarningsCoverRemainingBranches() {
    assertEquals(
        GridGrindProblemCode.CELL_NOT_FOUND,
        GridGrindProblems.codeFor(new CellNotFoundException("A1")));
    assertEquals(
        GridGrindProblemCode.MISSING_EXTERNAL_WORKBOOK,
        GridGrindProblems.codeFor(
            new MissingExternalWorkbookException(
                "Budget", "B4", "[Book2.xlsx]Sheet1!A1", "Book2.xlsx", "missing", null)));
    assertEquals(
        GridGrindProblemCode.UNREGISTERED_USER_DEFINED_FUNCTION,
        GridGrindProblems.codeFor(
            new UnregisteredUserDefinedFunctionException(
                "Budget", "B4", "FOO(A1)", "FOO", "missing udf", null)));
    assertEquals(List.of(), GridGrindProblems.causesFor(null));

    GridGrindResponse.ProblemContext.ExecuteStep enriched =
        assertInstanceOf(
            GridGrindResponse.ProblemContext.ExecuteStep.class,
            GridGrindProblems.enrichContext(
                new GridGrindResponse.ProblemContext.ExecuteStep(
                    "NEW", "NONE", 0, "step", "INSPECTION", null, null, null, null, null, null),
                new NamedRangeNotFoundException(
                    "LocalTotal", new ExcelNamedRangeScope.SheetScope("Budget"))));
    assertEquals("Budget", enriched.sheetName());
    assertEquals(
        "A1",
        assertInstanceOf(
                GridGrindResponse.ProblemContext.ExecuteStep.class,
                GridGrindProblems.enrichContext(
                    new GridGrindResponse.ProblemContext.ExecuteStep(
                        "CELL", "NONE", 1, "cell", "MUTATION", null, null, null, null, null, null),
                    new CellNotFoundException("A1")))
            .address());
    assertEquals(
        "BAD!",
        assertInstanceOf(
                GridGrindResponse.ProblemContext.ExecuteStep.class,
                GridGrindProblems.enrichContext(
                    new GridGrindResponse.ProblemContext.ExecuteStep(
                        "CELL", "NONE", 2, "cell", "MUTATION", null, null, null, null, null, null),
                    new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"))))
            .address());

    WorkbookPlan inspectionOnly =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(),
            List.of(
                inspect(
                    "workbook",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookSummary())));
    assertEquals(List.of(), GridGrindRequestWarnings.collect(inspectionOnly));
  }

  @Test
  void executorPolicyHelpersCoverExecutionModesPersistenceAndRuntimeGuards() throws IOException {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    assertExecutionModesAndValidation(executor);
    assertDeleteAndSourceHelpers();
    assertStreamingPersistenceBehaviors();
    assertRuntimeGuardBehaviors();
  }

  @Test
  void selectorAndActionDiagnosticsCoverTheRemainingExtractionFamilies() {
    assertActionDiagnostics();
    assertSelectorSheetAndAddressDiagnostics();
    assertNamedRangeSelectorDiagnostics();
    assertExceptionAndStepDiagnostics();
  }

  @Test
  void eventReadExecutionReturnsStructuredStepFailureForInspectionErrors() throws IOException {
    Path workbookPath = createWorkbookFile("gridgrind-event-step-failure-");

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        List.of(
                            inspect(
                                "missing-sheet",
                                new SheetSelector.ByName("Missing"),
                                new InspectionQuery.GetSheetSummary())))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals("Missing", failure.problem().context().sheetName());
  }

  @Test
  void eventReadExecutionReturnsInspectionResultsWhenAllStepsSucceed() throws IOException {
    Path workbookPath = createWorkbookFile("gridgrind-event-success-");

    GridGrindResponse.Success success =
        assertInstanceOf(
            GridGrindResponse.Success.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        List.of(
                            inspect(
                                "summary",
                                new WorkbookSelector.Current(),
                                new InspectionQuery.GetWorkbookSummary())))));

    assertEquals(1, success.inspections().size());
    assertEquals("summary", success.inspections().getFirst().stepId());
  }

  @Test
  void eventReadCloseFailureTurnsSuccessIntoExecuteRequestFailure() throws IOException {
    Path workbookPath = createWorkbookFile("gridgrind-event-close-success-failure-");

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class,
            new DefaultGridGrindRequestExecutor(
                    new WorkbookCommandExecutor(),
                    new WorkbookReadExecutor(),
                    ExcelWorkbook::close,
                    Files::createTempFile,
                    ignored -> {
                      throw new IOException("close failed");
                    })
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        List.of(
                            inspect(
                                "summary",
                                new WorkbookSelector.Current(),
                                new InspectionQuery.GetWorkbookSummary())))));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("EXECUTE_REQUEST", failure.problem().context().stage());
  }

  @Test
  void eventReadCloseFailureAppendsCauseToExistingStepFailure() throws IOException {
    Path workbookPath = createWorkbookFile("gridgrind-event-close-step-failure-");

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class,
            new DefaultGridGrindRequestExecutor(
                    new WorkbookCommandExecutor(),
                    new WorkbookReadExecutor(),
                    ExcelWorkbook::close,
                    Files::createTempFile,
                    ignored -> {
                      throw new IOException("close failed");
                    })
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        List.of(
                            inspect(
                                "missing-sheet",
                                new SheetSelector.ByName("Missing"),
                                new InspectionQuery.GetSheetSummary())))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals(2, failure.problem().causes().size());
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().causes().get(1).code());
  }

  @Test
  void eventReadExecutionReturnsOpenWorkbookFailureWhenMaterializationFails() {
    Path missingWorkbook = Path.of("/tmp/gridgrind-missing-event-" + System.nanoTime() + ".xlsx");

    GridGrindResponse.Failure failure =
        assertInstanceOf(
            GridGrindResponse.Failure.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(missingWorkbook.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        List.of(
                            inspect(
                                "summary",
                                new WorkbookSelector.Current(),
                                new InspectionQuery.GetWorkbookSummary())))));

    assertEquals(GridGrindProblemCode.WORKBOOK_NOT_FOUND, failure.problem().code());
    assertEquals("OPEN_WORKBOOK", failure.problem().context().stage());
  }

  private static void assertExecutionModesAndValidation(DefaultGridGrindRequestExecutor executor) {
    WorkbookPlan defaultModesRequest =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(),
            List.of());
    ExecutionModeSelection defaults =
        DefaultGridGrindRequestExecutor.executionModes(defaultModesRequest);
    assertEquals(ExecutionModeInput.ReadMode.FULL_XSSF, defaults.readMode());
    assertEquals(ExecutionModeInput.WriteMode.FULL_XSSF, defaults.writeMode());

    WorkbookPlan eventReadRequest =
        request(
            new WorkbookPlan.WorkbookSource.ExistingFile("/tmp/book.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None(),
            new ExecutionModeInput(
                ExecutionModeInput.ReadMode.EVENT_READ, ExecutionModeInput.WriteMode.FULL_XSSF),
            null,
            List.of(),
            List.of(
                inspect(
                    "workbook",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookSummary())));
    ExecutionModeSelection eventModes =
        DefaultGridGrindRequestExecutor.executionModes(eventReadRequest);
    assertTrue(
        DefaultGridGrindRequestExecutor.directEventReadEligible(eventReadRequest, eventModes));

    WorkbookPlan streamingInspectionBeforeEnsure =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            new ExecutionModeInput(
                ExecutionModeInput.ReadMode.FULL_XSSF,
                ExecutionModeInput.WriteMode.STREAMING_WRITE),
            null,
            List.of(),
            List.of(
                inspect(
                    "workbook",
                    new WorkbookSelector.Current(),
                    new InspectionQuery.GetWorkbookSummary())));
    Optional<String> executionModeFailure =
        executor.executionModeFailure(streamingInspectionBeforeEnsure);
    assertTrue(executionModeFailure.orElseThrow().contains("before any inspection step"));
  }

  private static void assertDeleteAndSourceHelpers() throws IOException {
    Path deleteTarget = Files.createTempFile("gridgrind-delete-", ".tmp");
    ExecutionWorkbookSupport.deleteIfExists(deleteTarget);
    assertFalse(Files.exists(deleteTarget));

    Path retained = Files.createTempFile("gridgrind-delete-retained-", ".tmp");
    ExecutionWorkbookSupport.deleteIfExists(
        retained,
        ignored -> {
          throw new IOException("best effort");
        });
    assertTrue(Files.exists(retained));
    ExecutionWorkbookSupport.deleteIfExists(retained);
    ExecutionWorkbookSupport.deleteIfExists(
        null,
        ignored -> {
          throw new AssertionError("null paths must not invoke the deleter");
        });

    ExcelOoxmlPersistenceOptions noneOptions =
        ExecutionRequestPaths.persistenceOptions(new WorkbookPlan.WorkbookPersistence.None());
    ExcelOoxmlPersistenceOptions saveAsOptions =
        ExecutionRequestPaths.persistenceOptions(
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "/tmp/out.xlsx",
                new OoxmlPersistenceSecurityInput(
                    new OoxmlEncryptionInput("secret", ExcelOoxmlEncryptionMode.AGILE), null)));
    ExcelOoxmlPersistenceOptions overwriteOptions =
        ExecutionRequestPaths.persistenceOptions(
            new WorkbookPlan.WorkbookPersistence.OverwriteSource(
                new OoxmlPersistenceSecurityInput(
                    new OoxmlEncryptionInput("secret", ExcelOoxmlEncryptionMode.AGILE), null)));
    assertTrue(noneOptions.isEmpty());
    assertFalse(saveAsOptions.isEmpty());
    assertFalse(overwriteOptions.isEmpty());

    ExcelOoxmlPackageSecuritySnapshot newSecurity =
        ExecutionRequestPaths.sourcePackageSecurity(new WorkbookPlan.WorkbookSource.New());
    ExcelOoxmlPackageSecuritySnapshot existingSecurity =
        ExecutionRequestPaths.sourcePackageSecurity(
            new WorkbookPlan.WorkbookSource.ExistingFile("/tmp/in.xlsx"));
    assertFalse(newSecurity.isSecure());
    assertFalse(existingSecurity.isSecure());
    assertNull(
        ExecutionRequestPaths.sourceEncryptionPassword(new WorkbookPlan.WorkbookSource.New()));
    assertNull(
        ExecutionRequestPaths.sourceEncryptionPassword(
            new WorkbookPlan.WorkbookSource.ExistingFile("/tmp/in.xlsx")));
    assertEquals(
        "open-secret",
        ExecutionRequestPaths.sourceEncryptionPassword(
            new WorkbookPlan.WorkbookSource.ExistingFile(
                "/tmp/in.xlsx", new OoxmlOpenSecurityInput("open-secret"))));
  }

  private static void assertStreamingPersistenceBehaviors() throws IOException {
    ExecutionWorkbookSupport workbookSupport = new ExecutionWorkbookSupport(Files::createTempFile);
    Path materialized = createWorkbookFile("gridgrind-streaming-source-");
    GridGrindResponse.PersistenceOutcome notSaved =
        workbookSupport.persistStreamingWorkbook(
            materialized,
            new WorkbookPlan.WorkbookPersistence.None(),
            new WorkbookPlan.WorkbookSource.New());
    assertInstanceOf(GridGrindResponse.PersistenceOutcome.NotSaved.class, notSaved);

    Path saveAsRoot = Files.createTempDirectory("gridgrind-streaming-saveas-");
    Path saveAsPath = saveAsRoot.resolve("nested output").resolve("streaming save-as.xlsx");
    GridGrindResponse.PersistenceOutcome.SavedAs savedAs =
        assertInstanceOf(
            GridGrindResponse.PersistenceOutcome.SavedAs.class,
            workbookSupport.persistStreamingWorkbook(
                materialized,
                new WorkbookPlan.WorkbookPersistence.SaveAs(saveAsPath.toString()),
                new WorkbookPlan.WorkbookSource.New()));
    assertEquals(saveAsPath.toAbsolutePath().toString(), savedAs.executionPath());
    assertTrue(Files.exists(saveAsPath));

    Path overwriteMaterialized = createWorkbookFile("gridgrind-streaming-overwrite-materialized-");
    Path overwriteSourcePath = createWorkbookFile("gridgrind-streaming-overwrite-source-");
    GridGrindResponse.PersistenceOutcome.Overwritten overwritten =
        assertInstanceOf(
            GridGrindResponse.PersistenceOutcome.Overwritten.class,
            workbookSupport.persistStreamingWorkbook(
                overwriteMaterialized,
                new WorkbookPlan.WorkbookPersistence.OverwriteSource(),
                new WorkbookPlan.WorkbookSource.ExistingFile(overwriteSourcePath.toString())));
    assertEquals(overwriteSourcePath.toAbsolutePath().toString(), overwritten.executionPath());

    IllegalArgumentException overwriteFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                workbookSupport.persistStreamingWorkbook(
                    materialized,
                    new WorkbookPlan.WorkbookPersistence.OverwriteSource(),
                    new WorkbookPlan.WorkbookSource.New()));
    assertEquals(
        "OVERWRITE persistence requires an EXISTING source", overwriteFailure.getMessage());
  }

  private static void assertRuntimeGuardBehaviors() throws IOException {
    ExecutionResponseSupport responseSupport =
        new ExecutionResponseSupport(
            ExcelWorkbook::close, ExcelOoxmlPackageSecuritySupport.ReadableWorkbook::close);
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(),
            List.of());
    GridGrindResponse.Failure runtimeFailure =
        assertInstanceOf(
            GridGrindResponse.Failure.class,
            responseSupport.guardUnexpectedRuntime(
                GridGrindProtocolVersion.V1,
                request,
                ExecutionJournalRecorder.start(request, ExecutionJournalSink.NOOP),
                () -> {
                  throw new UnsupportedOperationException("boom");
                }));
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, runtimeFailure.problem().code());

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      GridGrindResponse.Failure workbookRuntimeFailure =
          assertInstanceOf(
              GridGrindResponse.Failure.class,
              responseSupport.guardUnexpectedRuntime(
                  GridGrindProtocolVersion.V1,
                  request,
                  ExecutionJournalRecorder.start(request, ExecutionJournalSink.NOOP),
                  workbook,
                  () -> {
                    throw new UnsupportedOperationException("boom");
                  }));
      assertEquals(GridGrindProblemCode.INTERNAL_ERROR, workbookRuntimeFailure.problem().code());
    }

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExecutionResponseSupport closeFailingResponseSupport =
          new ExecutionResponseSupport(
              ignored -> {
                throw new IOException("close failed");
              },
              ExcelOoxmlPackageSecuritySupport.ReadableWorkbook::close);
      GridGrindResponse.Failure closeFailure =
          assertInstanceOf(
              GridGrindResponse.Failure.class,
              closeFailingResponseSupport.guardUnexpectedRuntime(
                  GridGrindProtocolVersion.V1,
                  request,
                  ExecutionJournalRecorder.start(request, ExecutionJournalSink.NOOP),
                  workbook,
                  () -> {
                    throw new UnsupportedOperationException("boom");
                  }));
      assertEquals(2, closeFailure.problem().causes().size());
      assertEquals(GridGrindProblemCode.IO_ERROR, closeFailure.problem().causes().get(1).code());
    }
  }

  private static void assertActionDiagnostics() {
    MutationAction.SetCell setBlank = new MutationAction.SetCell(new CellInput.Blank());
    MutationAction.SetCell setText = new MutationAction.SetCell(textCell("plain text"));
    MutationAction.SetCell setRichText =
        new MutationAction.SetCell(new CellInput.RichText(List.of(richTextRun("rich"))));
    MutationAction.SetCell setNumeric = new MutationAction.SetCell(new CellInput.Numeric(12.5));
    MutationAction.SetCell setBoolean =
        new MutationAction.SetCell(new CellInput.BooleanValue(true));
    MutationAction.SetCell setDate =
        new MutationAction.SetCell(new CellInput.Date(LocalDate.parse("2026-04-17")));
    MutationAction.SetCell setDateTime =
        new MutationAction.SetCell(
            new CellInput.DateTime(LocalDateTime.parse("2026-04-17T09:10:11")));
    MutationAction.SetCell setFormula = new MutationAction.SetCell(formulaCell("SUM(A1:A2)"));
    MutationAction.SetPivotTable pivotFromRange =
        pivotTableAction(new PivotTableInput.Source.Range("Budget", "A1:B5"));
    MutationAction.SetPivotTable pivotFromNamedRange =
        pivotTableAction(new PivotTableInput.Source.NamedRange("BudgetSource"));
    MutationAction.SetPivotTable pivotFromTable =
        pivotTableAction(new PivotTableInput.Source.Table("BudgetTable"));
    MutationAction.SetNamedRange sheetScopedNamedRange =
        new MutationAction.SetNamedRange(
            "LocalTotal",
            new NamedRangeScope.Sheet("Budget"),
            new NamedRangeTarget("SUM(Budget!B2:B4)"));
    MutationAction.SetNamedRange workbookScopedNamedRange =
        new MutationAction.SetNamedRange(
            "BudgetTotal",
            new NamedRangeScope.Workbook(),
            new NamedRangeTarget("SUM(Budget!B2:B4)"));
    MutationAction.SetConditionalFormatting singleRangeFormatting =
        conditionalFormattingAction(List.of("B2:B5"));
    MutationAction.SetConditionalFormatting multiRangeFormatting =
        conditionalFormattingAction(List.of("B2:B5", "D2:D5"));

    assertEquals("Budget", ExecutionDiagnosticFields.sheetNameFor(pivotFromRange));
    assertEquals("C5", ExecutionDiagnosticFields.addressFor(pivotFromRange));
    assertEquals("A1:B5", ExecutionDiagnosticFields.rangeFor(pivotFromRange));
    assertNull(ExecutionDiagnosticFields.rangeFor(pivotFromNamedRange));
    assertNull(ExecutionDiagnosticFields.rangeFor(pivotFromTable));
    assertEquals("BudgetSource", ExecutionDiagnosticFields.namedRangeNameFor(pivotFromNamedRange));
    assertNull(ExecutionDiagnosticFields.namedRangeNameFor(pivotFromRange));
    assertNull(ExecutionDiagnosticFields.namedRangeNameFor(pivotFromTable));
    assertEquals("Budget", ExecutionDiagnosticFields.sheetNameFor(sheetScopedNamedRange));
    assertNull(ExecutionDiagnosticFields.sheetNameFor(workbookScopedNamedRange));
    assertEquals("SUM(Budget!B2:B4)", ExecutionDiagnosticFields.formulaFor(sheetScopedNamedRange));
    assertNull(ExecutionDiagnosticFields.formulaFor(pivotFromRange));
    assertNull(ExecutionDiagnosticFields.sheetNameFor(setText));
    assertNull(ExecutionDiagnosticFields.formulaFor(setBlank));
    assertNull(ExecutionDiagnosticFields.formulaFor(setText));
    assertNull(ExecutionDiagnosticFields.formulaFor(setRichText));
    assertNull(ExecutionDiagnosticFields.formulaFor(setNumeric));
    assertNull(ExecutionDiagnosticFields.formulaFor(setBoolean));
    assertNull(ExecutionDiagnosticFields.formulaFor(setDate));
    assertNull(ExecutionDiagnosticFields.formulaFor(setDateTime));
    assertEquals("SUM(A1:A2)", ExecutionDiagnosticFields.formulaFor(setFormula));
    assertEquals("B2:B5", ExecutionDiagnosticFields.rangeFor(singleRangeFormatting));
    assertNull(ExecutionDiagnosticFields.rangeFor(multiRangeFormatting));
  }

  private static void assertSelectorSheetAndAddressDiagnostics() {
    Selector customSelector =
        () -> dev.erst.gridgrind.contract.selector.SelectorCardinality.ZERO_OR_ONE;

    assertEquals(
        "Budget",
        ExecutionDiagnosticFields.sheetNameFor(new DrawingObjectSelector.AllOnSheet("Budget")));
    assertEquals(
        "Budget",
        ExecutionDiagnosticFields.sheetNameFor(new DrawingObjectSelector.ByName("Budget", "Logo")));
    assertEquals(
        "Budget",
        ExecutionDiagnosticFields.sheetNameFor(new ChartSelector.ByName("Budget", "Revenue")));
    assertEquals(
        "Budget", ExecutionDiagnosticFields.sheetNameFor(new ChartSelector.AllOnSheet("Budget")));
    assertNull(ExecutionDiagnosticFields.sheetNameFor(new TableSelector.All()));
    assertNull(ExecutionDiagnosticFields.sheetNameFor(new TableSelector.ByName("BudgetTable")));
    assertNull(
        ExecutionDiagnosticFields.sheetNameFor(
            new TableSelector.ByNames(List.of("BudgetTable", "ForecastTable"))));
    assertEquals(
        "Budget",
        ExecutionDiagnosticFields.sheetNameFor(
            new TableSelector.ByNameOnSheet("BudgetTable", "Budget")));
    assertNull(ExecutionDiagnosticFields.sheetNameFor(new PivotTableSelector.All()));
    assertNull(ExecutionDiagnosticFields.sheetNameFor(new PivotTableSelector.ByName("SalesPivot")));
    assertNull(
        ExecutionDiagnosticFields.sheetNameFor(
            new PivotTableSelector.ByNames(List.of("SalesPivot", "ForecastPivot"))));
    assertEquals(
        "Budget",
        ExecutionDiagnosticFields.sheetNameFor(
            new PivotTableSelector.ByNameOnSheet("SalesPivot", "Budget")));
    assertNull(ExecutionDiagnosticFields.sheetNameFor(new NamedRangeSelector.All()));
    assertNull(
        ExecutionDiagnosticFields.sheetNameFor(new NamedRangeSelector.ByName("BudgetTotal")));
    assertNull(
        ExecutionDiagnosticFields.sheetNameFor(
            new NamedRangeSelector.ByNames(List.of("BudgetTotal", "ForecastTotal"))));
    assertNull(
        ExecutionDiagnosticFields.sheetNameFor(
            new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertEquals(
        "Budget",
        ExecutionDiagnosticFields.sheetNameFor(
            new NamedRangeSelector.SheetScope("BudgetTotal", "Budget")));
    assertEquals(
        "Budget",
        ExecutionDiagnosticFields.sheetNameFor(
            new TableRowSelector.AllRows(
                new TableSelector.ByNameOnSheet("BudgetTable", "Budget"))));
    assertEquals(
        "Budget",
        ExecutionDiagnosticFields.sheetNameFor(
            new TableRowSelector.ByIndex(
                new TableSelector.ByNameOnSheet("BudgetTable", "Budget"), 0)));
    assertEquals(
        "Budget",
        ExecutionDiagnosticFields.sheetNameFor(
            new TableCellSelector.ByColumnName(
                new TableRowSelector.ByKeyCell(
                    new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
                    "Item",
                    textCell("Hosting")),
                "Amount")));
    assertEquals(
        "Budget",
        ExecutionDiagnosticFields.sheetNameFor(new SheetSelector.ByNames(List.of("Budget"))));
    assertNull(
        ExecutionDiagnosticFields.sheetNameFor(
            new SheetSelector.ByNames(List.of("Budget", "Forecast"))));
    assertNull(ExecutionDiagnosticFields.sheetNameFor(customSelector));
    assertNull(ExecutionDiagnosticFields.addressFor(new SheetSelector.ByName("Budget")));
    assertNull(ExecutionDiagnosticFields.addressFor(customSelector));

    assertEquals(
        "A1",
        ExecutionDiagnosticFields.addressFor(
            new CellSelector.ByQualifiedAddresses(
                List.of(new CellSelector.QualifiedAddress("Budget", "A1")))));
    assertNull(
        ExecutionDiagnosticFields.addressFor(
            new CellSelector.ByQualifiedAddresses(
                List.of(
                    new CellSelector.QualifiedAddress("Budget", "A1"),
                    new CellSelector.QualifiedAddress("Forecast", "A1")))));
    assertEquals(
        "B2",
        ExecutionDiagnosticFields.addressFor(
            new RangeSelector.RectangularWindow("Budget", "B2", 2, 2)));
    assertEquals(
        "A1:B2",
        ExecutionDiagnosticFields.rangeFor(new RangeSelector.ByRanges("Budget", List.of("A1:B2"))));
    assertNull(
        ExecutionDiagnosticFields.rangeFor(
            new RangeSelector.ByRanges("Budget", List.of("A1:B2", "C1:D2"))));
    assertEquals(
        "B2:C3",
        ExecutionDiagnosticFields.rangeFor(
            new RangeSelector.RectangularWindow("Budget", "B2", 2, 2)));
  }

  private static void assertNamedRangeSelectorDiagnostics() {
    assertEquals(
        "Budget",
        ExecutionDiagnosticFields.singleSheetName(
            new CellSelector.ByQualifiedAddresses(
                List.of(new CellSelector.QualifiedAddress("Budget", "A1")))));
    assertNull(
        ExecutionDiagnosticFields.singleSheetName(
            new CellSelector.ByQualifiedAddresses(
                List.of(
                    new CellSelector.QualifiedAddress("Budget", "A1"),
                    new CellSelector.QualifiedAddress("Forecast", "A1")))));
    assertEquals(
        "Budget",
        ExecutionDiagnosticFields.singleSheetName(
            (NamedRangeSelector) new NamedRangeSelector.SheetScope("LocalTotal", "Budget")));
    assertEquals(
        "Budget",
        ExecutionDiagnosticFields.singleSheetName(
            (NamedRangeSelector)
                new NamedRangeSelector.AnyOf(
                    List.of(new NamedRangeSelector.SheetScope("LocalTotal", "Budget")))));
    assertNull(
        ExecutionDiagnosticFields.singleSheetName(
            new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.SheetScope("LocalTotal", "Budget"),
                    new NamedRangeSelector.SheetScope("ForecastTotal", "Forecast")))));
    assertNull(
        ExecutionDiagnosticFields.singleSheetName(
            (NamedRangeSelector) new NamedRangeSelector.All()));
    assertNull(
        ExecutionDiagnosticFields.singleSheetName(
            (NamedRangeSelector) new NamedRangeSelector.ByName("BudgetTotal")));
    assertNull(
        ExecutionDiagnosticFields.singleSheetName(
            (NamedRangeSelector) new NamedRangeSelector.ByNames(List.of("BudgetTotal"))));
    assertNull(
        ExecutionDiagnosticFields.singleSheetName(
            (NamedRangeSelector) new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertNull(
        ExecutionDiagnosticFields.singleSheetName(
            (NamedRangeSelector.Ref) new NamedRangeSelector.ByName("BudgetTotal")));
    assertNull(
        ExecutionDiagnosticFields.singleSheetName(
            (NamedRangeSelector.Ref) new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertEquals(
        "Budget",
        ExecutionDiagnosticFields.singleSheetName(
            (NamedRangeSelector.Ref) new NamedRangeSelector.SheetScope("BudgetTotal", "Budget")));

    assertNull(ExecutionDiagnosticFields.singleNamedRangeName(new NamedRangeSelector.All()));
    assertEquals(
        "BudgetTotal",
        ExecutionDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector) new NamedRangeSelector.ByName("BudgetTotal")));
    assertEquals(
        "BudgetTotal",
        ExecutionDiagnosticFields.singleNamedRangeName(
            new NamedRangeSelector.ByNames(List.of("BudgetTotal"))));
    assertNull(
        ExecutionDiagnosticFields.singleNamedRangeName(
            new NamedRangeSelector.ByNames(List.of("BudgetTotal", "ForecastTotal"))));
    assertEquals(
        "BudgetTotal",
        ExecutionDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector) new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertEquals(
        "BudgetTotal",
        ExecutionDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector) new NamedRangeSelector.SheetScope("BudgetTotal", "Budget")));
    assertEquals(
        "BudgetTotal",
        ExecutionDiagnosticFields.singleNamedRangeName(
            new NamedRangeSelector.AnyOf(
                List.of(new NamedRangeSelector.WorkbookScope("BudgetTotal")))));
    assertNull(
        ExecutionDiagnosticFields.singleNamedRangeName(
            new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                    new NamedRangeSelector.WorkbookScope("ForecastTotal")))));
    assertEquals(
        "BudgetTotal",
        ExecutionDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector.Ref) new NamedRangeSelector.ByName("BudgetTotal")));
    assertEquals(
        "BudgetTotal",
        ExecutionDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector.Ref) new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertEquals(
        "BudgetTotal",
        ExecutionDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector.Ref) new NamedRangeSelector.SheetScope("BudgetTotal", "Budget")));
  }

  private static void assertExceptionAndStepDiagnostics() {
    MutationAction.SetCell setText = new MutationAction.SetCell(textCell("plain text"));
    MutationAction.SetPivotTable pivotFromRange =
        pivotTableAction(new PivotTableInput.Source.Range("Budget", "A1:B5"));
    MutationStep pivotStep =
        new MutationStep(
            "pivot-step",
            new PivotTableSelector.ByNameOnSheet("SalesPivot", "Budget"),
            pivotFromRange);
    MutationStep addressedCellStep =
        new MutationStep("cell-step", new CellSelector.ByAddress("Budget", "D4"), setText);

    assertEquals(
        "Budget", ExecutionDiagnosticFields.sheetNameFor(new SheetNotFoundException("Budget")));
    assertNull(ExecutionDiagnosticFields.sheetNameFor(new RuntimeException("x")));
    assertNull(
        ExecutionDiagnosticFields.sheetNameFor(
            new NamedRangeNotFoundException(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
    assertEquals(
        "Budget",
        ExecutionDiagnosticFields.sheetNameFor(
            new NamedRangeNotFoundException(
                "LocalTotal", new ExcelNamedRangeScope.SheetScope("Budget"))));
    assertEquals("A1", ExecutionDiagnosticFields.addressFor(new CellNotFoundException("A1")));
    assertEquals(
        "BAD!",
        ExecutionDiagnosticFields.addressFor(
            new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"))));
    assertNull(ExecutionDiagnosticFields.addressFor(new RuntimeException("x")));
    assertEquals("C5", ExecutionDiagnosticFields.addressFor(pivotStep));
    assertEquals("D4", ExecutionDiagnosticFields.addressFor(addressedCellStep));
  }

  private static MutationAction.SetPivotTable pivotTableAction(PivotTableInput.Source source) {
    return new MutationAction.SetPivotTable(
        new PivotTableInput(
            "SalesPivot",
            "Budget",
            source,
            new PivotTableInput.Anchor("C5"),
            List.of(),
            List.of(),
            List.of(),
            List.of(
                new PivotTableInput.DataField(
                    "Amount",
                    dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction.SUM,
                    null,
                    null))));
  }

  private static MutationAction.SetConditionalFormatting conditionalFormattingAction(
      List<String> ranges) {
    return new MutationAction.SetConditionalFormatting(
        new ConditionalFormattingBlockInput(
            ranges,
            List.of(
                new ConditionalFormattingRuleInput.FormulaRule(
                    "B2>0",
                    true,
                    new DifferentialStyleInput(
                        null, true, null, null, null, null, null, null, null)))));
  }

  private static Path createWorkbookFile(String prefix) throws IOException {
    Path workbookPath = Files.createTempFile(prefix, ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").setCell("A1", ExcelCellValue.text("Header"));
      workbook.save(workbookPath);
    }
    return workbookPath;
  }
}
