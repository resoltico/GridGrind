package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingDefinitionInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.DifferentialStyleInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence;
import dev.erst.gridgrind.contract.query.*;
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
import dev.erst.gridgrind.contract.step.WorkbookStaticRequestContract;
import dev.erst.gridgrind.contract.step.WorkbookStaticViolation;
import dev.erst.gridgrind.excel.CellNotFoundException;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.MissingExternalWorkbookException;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.UnregisteredUserDefinedFunctionException;
import dev.erst.gridgrind.excel.WorkbookArtifactIo;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.excel.foundation.ExcelOoxmlWriteHash;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackageSecuritySnapshot;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceOptions;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Coverage for executor policy seams, helper extraction, and failure routing. */
class ExecutorPolicyCoverageTest {
  private static List<String> staticValidationMessages(WorkbookPlan request) {
    return WorkbookStaticRequestContract.validate(WorkbookStaticRequestContract.from(request))
        .stream()
        .map(WorkbookStaticViolation::message)
        .toList();
  }

  @Test
  void inspectionCommandConverterRejectsUnsupportedChartTargets() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                InspectionCommandConverter.toReadCommand(
                    "charts",
                    new WorkbookSelector.Current(),
                    new WorkbookAssetIntrospectionQuery.GetCharts()));

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
                    new StructuredMutationAction.SetTable(
                        TableInput.withDefaultMetadata(
                            "BudgetTable", "Budget", "A1:B2", false, new TableStyleInput.None()))));
    assertEquals(
        "SET_TABLE target must match table.name and table.sheetName", tableMismatch.getMessage());

    IllegalArgumentException tableSheetMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookCommandConverter.toCommand(
                    new TableSelector.ByNameOnSheet("BudgetTable", "Archive"),
                    new StructuredMutationAction.SetTable(
                        TableInput.withDefaultMetadata(
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
                    new StructuredMutationAction.SetPivotTable(
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
                                    "Amount",
                                    Optional.empty()))))));
    assertEquals(
        "SET_PIVOT_TABLE target must match pivotTable.name and pivotTable.sheetName",
        pivotMismatch.getMessage());

    IllegalArgumentException pivotSheetMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookCommandConverter.toCommand(
                    new PivotTableSelector.ByNameOnSheet("SalesPivot", "Archive"),
                    new StructuredMutationAction.SetPivotTable(
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
                                    "Amount",
                                    Optional.empty()))))));
    assertEquals(
        "SET_PIVOT_TABLE target must match pivotTable.name and pivotTable.sheetName",
        pivotSheetMismatch.getMessage());

    IllegalArgumentException namedRangeMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookCommandConverter.toCommand(
                    new NamedRangeSelector.SheetScope("BudgetTotal", "Archive"),
                    new StructuredMutationAction.SetNamedRange(
                        "BudgetTotal",
                        new NamedRangeScope.Sheet("Budget"),
                        NamedRangeTarget.range("Budget", "B4"))));
    assertEquals(
        "SET_NAMED_RANGE target must match action name and scope", namedRangeMismatch.getMessage());

    IllegalArgumentException namedRangeNameMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookCommandConverter.toCommand(
                    new NamedRangeSelector.SheetScope("OtherTotal", "Budget"),
                    new StructuredMutationAction.SetNamedRange(
                        "BudgetTotal",
                        new NamedRangeScope.Sheet("Budget"),
                        NamedRangeTarget.range("Budget", "B4"))));
    assertEquals(
        "SET_NAMED_RANGE target must match action name and scope",
        namedRangeNameMismatch.getMessage());

    IllegalArgumentException workbookScopeMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookCommandConverter.toCommand(
                    new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                    new StructuredMutationAction.SetNamedRange(
                        "BudgetTotal",
                        new NamedRangeScope.Sheet("Budget"),
                        NamedRangeTarget.range("Budget", "B4"))));
    assertEquals(
        "SET_NAMED_RANGE target must match action name and scope",
        workbookScopeMismatch.getMessage());

    IllegalArgumentException broadDelete =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                WorkbookCommandConverter.toCommand(
                    new NamedRangeSelector.ByName("BudgetTotal"),
                    new StructuredMutationAction.DeleteNamedRange()));
    assertEquals(
        "DELETE_NAMED_RANGE requires target type NAMED_RANGE_WORKBOOK_SCOPE or"
            + " NAMED_RANGE_SHEET_SCOPE but got NAMED_RANGE_BY_NAME",
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

    dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep enriched =
        assertInstanceOf(
            dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep.class,
            GridGrindProblems.enrichContext(
                new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep(
                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape
                        .known("NEW", "NONE"),
                    new dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces
                        .StepReference(0, "step", "INSPECTION", "GET_NAMED_RANGES"),
                    dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.ProblemLocation
                        .unknown()),
                new NamedRangeNotFoundException(
                    "LocalTotal", new ExcelNamedRangeScope.SheetScope("Budget"))));
    assertEquals(java.util.Optional.of("Budget"), enriched.sheetName());
    assertEquals(
        java.util.Optional.of("A1"),
        assertInstanceOf(
                dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep.class,
                GridGrindProblems.enrichContext(
                    new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep(
                        dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape
                            .known("NEW", "NONE"),
                        new dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces
                            .StepReference(1, "cell", "MUTATION", "SET_CELL"),
                        dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces
                            .ProblemLocation.unknown()),
                    new CellNotFoundException("A1")))
            .address());
    assertEquals(
        java.util.Optional.of("BAD!"),
        assertInstanceOf(
                dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep.class,
                GridGrindProblems.enrichContext(
                    new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep(
                        dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape
                            .known("NEW", "NONE"),
                        new dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces
                            .StepReference(2, "cell", "MUTATION", "SET_CELL"),
                        dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces
                            .ProblemLocation.unknown()),
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
                    new WorkbookIntrospectionQuery.GetWorkbookSummary())));
    assertEquals(List.of(), GridGrindRequestWarnings.collect(inspectionOnly));
  }

  @Test
  void executorPolicyHelpersCoverExecutionModesPersistenceAndRuntimeGuards() throws IOException {
    assertExecutionModesAndValidation();
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

  private static void assertExecutionModesAndValidation() {
    WorkbookPlan defaultModesRequest =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(),
            List.of());
    ExecutionModeInput defaults =
        DefaultGridGrindRequestExecutor.executionMode(defaultModesRequest);
    assertInstanceOf(ExecutionModeInput.FullXssf.class, defaults);

    WorkbookPlan eventReadRequest =
        request(
            new WorkbookPlan.WorkbookSource.ExistingFile("/tmp/book.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionModeInput.eventRead(),
            null,
            List.of(),
            List.of(
                inspect(
                    "workbook",
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.GetWorkbookSummary())));
    ExecutionModeInput eventModes = DefaultGridGrindRequestExecutor.executionMode(eventReadRequest);
    assertTrue(
        DefaultGridGrindRequestExecutor.directEventReadEligible(eventReadRequest, eventModes));

    WorkbookPlan streamingInspectionBeforeEnsure =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionModeInput.streamingWrite(),
            null,
            List.of(),
            List.of(
                inspect(
                    "workbook",
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.GetWorkbookSummary())));
    List<String> executionModeFailures = staticValidationMessages(streamingInspectionBeforeEnsure);
    assertTrue(
        executionModeFailures.stream()
            .anyMatch(failure -> failure.contains("before any inspection step")));
  }

  private static void assertDeleteAndSourceHelpers() throws IOException {
    Path workingDirectory = Path.of("/tmp/gridgrind-policy");
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

    try (var prepared = ExecutionInputBindingsFixtureSupport.preparedBindings(workingDirectory)) {
      ExcelOoxmlPersistenceOptions noneOptions =
          ExecutionRequestPaths.persistenceOptions(
              new WorkbookPlan.WorkbookPersistence.None(), prepared.bindings());
      ExcelOoxmlPersistenceOptions saveAsOptions =
          ExecutionRequestPaths.persistenceOptions(
              new WorkbookPlan.WorkbookPersistence.SaveAs(
                  "/tmp/out.xlsx",
                  WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                  new OoxmlPersistenceSecurityInput(
                      new dev.erst.gridgrind.contract.dto.OoxmlPersistenceEncryptionInput.Encrypt(
                          new OoxmlEncryptionInput(
                              "secret",
                              ExcelOoxmlWriteCipher.AES_256,
                              ExcelOoxmlWriteHash.SHA_512)),
                      new dev.erst.gridgrind.contract.dto.OoxmlPersistenceSignatureInput.None())),
              prepared.bindings());
      ExcelOoxmlPersistenceOptions overwriteOptions =
          ExecutionRequestPaths.persistenceOptions(
              new WorkbookPlan.WorkbookPersistence.Overwrite(
                  new OoxmlPersistenceSecurityInput(
                      new dev.erst.gridgrind.contract.dto.OoxmlPersistenceEncryptionInput.Encrypt(
                          new OoxmlEncryptionInput(
                              "secret",
                              ExcelOoxmlWriteCipher.AES_256,
                              ExcelOoxmlWriteHash.SHA_512)),
                      new dev.erst.gridgrind.contract.dto.OoxmlPersistenceSignatureInput.None())),
              prepared.bindings());
      assertTrue(noneOptions.writesPlaintextUnsigned());
      assertFalse(saveAsOptions.writesPlaintextUnsigned());
      assertFalse(overwriteOptions.writesPlaintextUnsigned());
    }

    ExcelOoxmlPackageSecuritySnapshot newSecurity =
        ExecutionRequestPaths.sourcePackageSecurity(new WorkbookPlan.WorkbookSource.New());
    ExcelOoxmlPackageSecuritySnapshot existingSecurity =
        ExecutionRequestPaths.sourcePackageSecurity(
            new WorkbookPlan.WorkbookSource.ExistingFile("/tmp/in.xlsx"));
    assertFalse(newSecurity.isSecure());
    assertFalse(existingSecurity.isSecure());
    assertEquals(
        Optional.empty(),
        ExecutionRequestPaths.sourceEncryptionPassword(new WorkbookPlan.WorkbookSource.New()));
    assertEquals(
        Optional.empty(),
        ExecutionRequestPaths.sourceEncryptionPassword(
            new WorkbookPlan.WorkbookSource.ExistingFile("/tmp/in.xlsx")));
    assertEquals(
        Optional.of("open-secret"),
        ExecutionRequestPaths.sourceEncryptionPassword(
            new WorkbookPlan.WorkbookSource.ExistingFile(
                "/tmp/in.xlsx", new OoxmlOpenSecurityInput(java.util.Optional.of("open-secret")))));
  }

  private static void assertStreamingPersistenceBehaviors() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-streaming-policy-");
    ExecutionWorkbookSupport workbookSupport =
        ExecutionContextFixtureSupport.workbookSupport(workingDirectory);
    Path materialized = createWorkbookFile("gridgrind-streaming-source-");
    try (var prepared = ExecutionInputBindingsFixtureSupport.preparedBindings(workingDirectory)) {
      WorkbookResultPersistence.PersistenceOutcome notSaved =
          workbookSupport.persistStreamingWorkbook(
              materialized,
              new WorkbookPlan.WorkbookPersistence.None(),
              new WorkbookPlan.WorkbookSource.New(),
              prepared.bindings());
      assertInstanceOf(WorkbookResultPersistence.PersistenceOutcome.NotSaved.class, notSaved);
    }

    Path outputDirectory = Files.createDirectory(workingDirectory.resolve("output"));
    Path saveAsPath = outputDirectory.resolve("streaming-save-as.xlsx");
    try (var prepared = ExecutionInputBindingsFixtureSupport.preparedBindings(workingDirectory)) {
      prepared
          .access()
          .prepareOutput(
              saveAsPath.toString(),
              "persistence",
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.CREATE_NEW);
      WorkbookResultPersistence.PersistenceOutcome.SavedAs savedAs =
          assertInstanceOf(
              WorkbookResultPersistence.PersistenceOutcome.SavedAs.class,
              workbookSupport.persistStreamingWorkbook(
                  materialized,
                  new WorkbookPlan.WorkbookPersistence.SaveAs(
                      saveAsPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                  new WorkbookPlan.WorkbookSource.New(),
                  prepared.bindings()));
      assertEquals(
          saveAsPath.toAbsolutePath().toString(),
          DefaultGridGrindRequestExecutorTestSupport.writtenExecutionPath(savedAs));
      assertTrue(Files.exists(saveAsPath));
    }

    Path overwriteMaterialized = createWorkbookFile("gridgrind-streaming-overwrite-materialized-");
    Path overwriteSourcePath = Files.createTempFile(workingDirectory, "overwrite-source-", ".xlsx");
    Files.copy(
        createWorkbookFile("gridgrind-streaming-overwrite-source-"),
        overwriteSourcePath,
        java.nio.file.StandardCopyOption.REPLACE_EXISTING);
    try (var prepared = ExecutionInputBindingsFixtureSupport.preparedBindings(workingDirectory)) {
      prepared
          .access()
          .prepareOutput(
              overwriteSourcePath.toString(),
              "persistence",
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING);
      WorkbookResultPersistence.PersistenceOutcome.Overwritten overwritten =
          assertInstanceOf(
              WorkbookResultPersistence.PersistenceOutcome.Overwritten.class,
              workbookSupport.persistStreamingWorkbook(
                  overwriteMaterialized,
                  new WorkbookPlan.WorkbookPersistence.Overwrite(
                      OoxmlPersistenceSecurityInput.none()),
                  new WorkbookPlan.WorkbookSource.ExistingFile(overwriteSourcePath.toString()),
                  prepared.bindings()));
      assertEquals(
          overwriteSourcePath.toAbsolutePath().toString(),
          DefaultGridGrindRequestExecutorTestSupport.writtenExecutionPath(overwritten));
    }

    try (var prepared = ExecutionInputBindingsFixtureSupport.preparedBindings(workingDirectory)) {
      IllegalArgumentException overwriteFailure =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  workbookSupport.persistStreamingWorkbook(
                      materialized,
                      new WorkbookPlan.WorkbookPersistence.Overwrite(
                          OoxmlPersistenceSecurityInput.none()),
                      new WorkbookPlan.WorkbookSource.New(),
                      prepared.bindings()));
      assertEquals(
          "OVERWRITE persistence requires an EXISTING source", overwriteFailure.getMessage());
    }
  }

  private static void assertRuntimeGuardBehaviors() throws IOException {
    ExecutionResponseSupport responseSupport =
        new ExecutionResponseSupport(
            ExcelWorkbook::close, WorkbookArtifactIo.MaterializedWorkbook::close);
    WorkbookPlan request =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            List.of(),
            List.of());
    WorkbookResult.Failure runtimeFailure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            responseSupport.guardUnexpectedRuntime(
                GridGrindProtocolVersion.V2,
                request,
                ExecutionContextFixtureSupport.startJournal(request, ExecutionProgressSink.NOOP),
                () -> {
                  throw new UnsupportedOperationException("boom");
                }));
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, runtimeFailure.problem().code());

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookResult.Failure workbookRuntimeFailure =
          assertInstanceOf(
              WorkbookResult.Failure.class,
              responseSupport.guardUnexpectedRuntime(
                  GridGrindProtocolVersion.V2,
                  request,
                  ExecutionContextFixtureSupport.startJournal(request, ExecutionProgressSink.NOOP),
                  workbook,
                  () -> {
                    throw new UnsupportedOperationException("boom");
                  }));
      assertEquals(GridGrindProblemCode.INTERNAL_ERROR, workbookRuntimeFailure.problem().code());
    }

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      ExecutionResponseSupport closeFailingResponseSupport =
          new ExecutionResponseSupport(
              ignored -> {
                throw new IOException("close failed");
              },
              WorkbookArtifactIo.MaterializedWorkbook::close);
      WorkbookResult.Failure closeFailure =
          assertInstanceOf(
              WorkbookResult.Failure.class,
              closeFailingResponseSupport.guardUnexpectedRuntime(
                  GridGrindProtocolVersion.V2,
                  request,
                  ExecutionContextFixtureSupport.startJournal(request, ExecutionProgressSink.NOOP),
                  workbook,
                  () -> {
                    throw new UnsupportedOperationException("boom");
                  }));
      assertEquals(2, closeFailure.problem().causes().size());
      assertEquals(GridGrindProblemCode.IO_ERROR, closeFailure.problem().causes().get(1).code());
    }
  }

  private static void assertActionDiagnostics() {
    CellMutationAction.SetCell setBlank = new CellMutationAction.SetCell(new CellInput.Blank());
    CellMutationAction.SetCell setText = new CellMutationAction.SetCell(textCell("plain text"));
    CellMutationAction.SetCell setRichText =
        new CellMutationAction.SetCell(new CellInput.RichText(List.of(richTextRun("rich"))));
    CellMutationAction.SetCell setNumeric =
        new CellMutationAction.SetCell(new CellInput.NumberValue(12.5));
    CellMutationAction.SetCell setBoolean =
        new CellMutationAction.SetCell(new CellInput.BooleanValue(true));
    CellMutationAction.SetCell setDate =
        new CellMutationAction.SetCell(new CellInput.Date(LocalDate.parse("2026-04-17")));
    CellMutationAction.SetCell setDateTime =
        new CellMutationAction.SetCell(
            new CellInput.DateTime(LocalDateTime.parse("2026-04-17T09:10:11")));
    CellMutationAction.SetCell setFormula =
        new CellMutationAction.SetCell(formulaCell("SUM(A1:A2)"));
    StructuredMutationAction.SetPivotTable pivotFromRange =
        pivotTableAction(new PivotTableInput.Source.Range("Budget", "A1:B5"));
    StructuredMutationAction.SetPivotTable pivotFromNamedRange =
        pivotTableAction(new PivotTableInput.Source.NamedRange("BudgetSource"));
    StructuredMutationAction.SetPivotTable pivotFromTable =
        pivotTableAction(new PivotTableInput.Source.Table("BudgetTable"));
    StructuredMutationAction.SetNamedRange sheetScopedNamedRange =
        new StructuredMutationAction.SetNamedRange(
            "LocalTotal",
            new NamedRangeScope.Sheet("Budget"),
            NamedRangeTarget.formula("SUM(Budget!B2:B4)"));
    StructuredMutationAction.SetNamedRange workbookScopedNamedRange =
        new StructuredMutationAction.SetNamedRange(
            "BudgetTotal",
            new NamedRangeScope.Workbook(),
            NamedRangeTarget.formula("SUM(Budget!B2:B4)"));
    StructuredMutationAction.SetConditionalFormatting conditionalFormatting =
        conditionalFormattingAction();

    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionActionDiagnosticFields.sheetNameFor(pivotFromRange));
    assertEquals(
        java.util.Optional.of("C5"), ExecutionActionDiagnosticFields.addressFor(pivotFromRange));
    assertEquals(
        java.util.Optional.of("A1:B5"), ExecutionActionDiagnosticFields.rangeFor(pivotFromRange));
    assertEquals(
        java.util.Optional.empty(), ExecutionActionDiagnosticFields.rangeFor(pivotFromNamedRange));
    assertEquals(
        java.util.Optional.empty(), ExecutionActionDiagnosticFields.rangeFor(pivotFromTable));
    assertEquals(
        java.util.Optional.of("BudgetSource"),
        ExecutionActionDiagnosticFields.namedRangeNameFor(pivotFromNamedRange));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionActionDiagnosticFields.namedRangeNameFor(pivotFromRange));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionActionDiagnosticFields.namedRangeNameFor(pivotFromTable));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionActionDiagnosticFields.sheetNameFor(sheetScopedNamedRange));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionActionDiagnosticFields.sheetNameFor(workbookScopedNamedRange));
    assertEquals(
        java.util.Optional.of("SUM(Budget!B2:B4)"),
        ExecutionActionDiagnosticFields.formulaFor(sheetScopedNamedRange));
    assertEquals(
        java.util.Optional.empty(), ExecutionActionDiagnosticFields.formulaFor(pivotFromRange));
    assertEquals(java.util.Optional.empty(), ExecutionActionDiagnosticFields.sheetNameFor(setText));
    assertEquals(java.util.Optional.empty(), ExecutionActionDiagnosticFields.formulaFor(setBlank));
    assertEquals(java.util.Optional.empty(), ExecutionActionDiagnosticFields.formulaFor(setText));
    assertEquals(
        java.util.Optional.empty(), ExecutionActionDiagnosticFields.formulaFor(setRichText));
    assertEquals(
        java.util.Optional.empty(), ExecutionActionDiagnosticFields.formulaFor(setNumeric));
    assertEquals(
        java.util.Optional.empty(), ExecutionActionDiagnosticFields.formulaFor(setBoolean));
    assertEquals(java.util.Optional.empty(), ExecutionActionDiagnosticFields.formulaFor(setDate));
    assertEquals(
        java.util.Optional.empty(), ExecutionActionDiagnosticFields.formulaFor(setDateTime));
    assertEquals(
        java.util.Optional.of("SUM(A1:A2)"),
        ExecutionActionDiagnosticFields.formulaFor(setFormula));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionActionDiagnosticFields.rangeFor(conditionalFormatting));
  }

  private static void assertSelectorSheetAndAddressDiagnostics() {
    Selector selectorWithoutSheetOrAddress = new WorkbookSelector.Current();

    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new DrawingObjectSelector.AllOnSheet("Budget")));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new DrawingObjectSelector.ByName("Budget", "Logo")));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new ChartSelector.ByName("Budget", "Revenue")));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(new ChartSelector.AllOnSheet("Budget")));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(new TableSelector.All()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(new TableSelector.ByName("BudgetTable")));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new TableSelector.ByNames(List.of("BudgetTable", "ForecastTable"))));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new TableSelector.ByNameOnSheet("BudgetTable", "Budget")));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(new PivotTableSelector.All()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new PivotTableSelector.ByName("SalesPivot")));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new PivotTableSelector.ByNames(List.of("SalesPivot", "ForecastPivot"))));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new PivotTableSelector.ByNameOnSheet("SalesPivot", "Budget")));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(new NamedRangeSelector.All()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new NamedRangeSelector.ByName("BudgetTotal")));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new NamedRangeSelector.ByNames(List.of("BudgetTotal", "ForecastTotal"))));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new NamedRangeSelector.SheetScope("BudgetTotal", "Budget")));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new TableRowSelector.AllRows(
                new TableSelector.ByNameOnSheet("BudgetTable", "Budget"))));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new TableRowSelector.ByIndex(
                new TableSelector.ByNameOnSheet("BudgetTable", "Budget"), 0)));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new TableCellSelector.ByColumnName(
                new TableRowSelector.ByKeyCell(
                    new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
                    "Item",
                    textCell("Hosting")),
                "Amount")));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new SheetSelector.ByNames(List.of("Budget"))));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new SheetSelector.ByNames(List.of("Budget", "Forecast"))));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(selectorWithoutSheetOrAddress));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.addressFor(new SheetSelector.ByName("Budget")));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.addressFor(selectorWithoutSheetOrAddress));

    assertEquals(
        java.util.Optional.of("B2"),
        ExecutionSelectorDiagnosticFields.addressFor(
            new RangeSelector.RectangularWindow("Budget", "B2", 2, 2)));
    assertEquals(
        java.util.Optional.of("A1:B2"),
        ExecutionSelectorDiagnosticFields.rangeFor(
            new RangeSelector.ByRanges("Budget", List.of("A1:B2"))));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.rangeFor(
            new RangeSelector.ByRanges("Budget", List.of("A1:B2", "C1:D2"))));
    assertEquals(
        java.util.Optional.of("B2:C3"),
        ExecutionSelectorDiagnosticFields.rangeFor(
            new RangeSelector.RectangularWindow("Budget", "B2", 2, 2)));
  }

  private static void assertNamedRangeSelectorDiagnostics() {
    assertEquals(
        Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector) new NamedRangeSelector.SheetScope("LocalTotal", "Budget")));
    assertEquals(
        Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector)
                new NamedRangeSelector.AnyOf(
                    List.of(new NamedRangeSelector.SheetScope("LocalTotal", "Budget")))));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.SheetScope("LocalTotal", "Budget"),
                    new NamedRangeSelector.SheetScope("ForecastTotal", "Forecast")))));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector) new NamedRangeSelector.All()));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector) new NamedRangeSelector.ByName("BudgetTotal")));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector) new NamedRangeSelector.ByNames(List.of("BudgetTotal"))));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector) new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector.Ref) new NamedRangeSelector.ByName("BudgetTotal")));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector.Ref) new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertEquals(
        Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector.Ref) new NamedRangeSelector.SheetScope("BudgetTotal", "Budget")));

    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(new NamedRangeSelector.All()));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector) new NamedRangeSelector.ByName("BudgetTotal")));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            new NamedRangeSelector.ByNames(List.of("BudgetTotal"))));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            new NamedRangeSelector.ByNames(List.of("BudgetTotal", "ForecastTotal"))));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector) new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector) new NamedRangeSelector.SheetScope("BudgetTotal", "Budget")));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            new NamedRangeSelector.AnyOf(
                List.of(new NamedRangeSelector.WorkbookScope("BudgetTotal")))));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                    new NamedRangeSelector.WorkbookScope("ForecastTotal")))));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector.Ref) new NamedRangeSelector.ByName("BudgetTotal")));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector.Ref) new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector.Ref) new NamedRangeSelector.SheetScope("BudgetTotal", "Budget")));
  }

  private static void assertExceptionAndStepDiagnostics() {
    CellMutationAction.SetCell setText = new CellMutationAction.SetCell(textCell("plain text"));
    StructuredMutationAction.SetPivotTable pivotFromRange =
        pivotTableAction(new PivotTableInput.Source.Range("Budget", "A1:B5"));
    MutationStep pivotStep =
        new MutationStep(
            "pivot-step",
            new PivotTableSelector.ByNameOnSheet("SalesPivot", "Budget"),
            pivotFromRange);
    MutationStep addressedCellStep =
        new MutationStep("cell-step", new CellSelector.ByAddress("Budget", "D4"), setText);

    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionExceptionDiagnosticFields.sheetNameFor(new SheetNotFoundException("Budget")));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionExceptionDiagnosticFields.sheetNameFor(new RuntimeException("x")));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionExceptionDiagnosticFields.sheetNameFor(
            new NamedRangeNotFoundException(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionExceptionDiagnosticFields.sheetNameFor(
            new NamedRangeNotFoundException(
                "LocalTotal", new ExcelNamedRangeScope.SheetScope("Budget"))));
    assertEquals(
        java.util.Optional.of("A1"),
        ExecutionExceptionDiagnosticFields.addressFor(new CellNotFoundException("A1")));
    assertEquals(
        java.util.Optional.of("BAD!"),
        ExecutionExceptionDiagnosticFields.addressFor(
            new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"))));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionExceptionDiagnosticFields.addressFor(new RuntimeException("x")));
    assertEquals(java.util.Optional.of("C5"), ExecutionDiagnosticFields.addressFor(pivotStep));
    assertEquals(
        java.util.Optional.of("D4"), ExecutionDiagnosticFields.addressFor(addressedCellStep));
  }

  private static StructuredMutationAction.SetPivotTable pivotTableAction(
      PivotTableInput.Source source) {
    return new StructuredMutationAction.SetPivotTable(
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
                    "Amount",
                    Optional.empty()))));
  }

  private static StructuredMutationAction.SetConditionalFormatting conditionalFormattingAction() {
    return new StructuredMutationAction.SetConditionalFormatting(
        new ConditionalFormattingDefinitionInput(
            List.of(
                new ConditionalFormattingRuleInput.FormulaRule(
                    "B2>0",
                    true,
                    java.util.Optional.of(
                        new DifferentialStyleInput(
                            Optional.empty(),
                            Optional.of(true),
                            Optional.empty(),
                            Optional.empty(),
                            java.util.Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            java.util.Optional.empty(),
                            java.util.Optional.empty()))))));
  }
}
