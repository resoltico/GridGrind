package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.pivot.PivotTableSpec;
import dev.erst.gridgrind.core.model.security.OoxmlEncryptionSpec;
import dev.erst.gridgrind.core.model.security.OoxmlOpenSecuritySpec;
import dev.erst.gridgrind.core.model.security.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceSecurity;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingRuleSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.result.WorkbookResultPersistence;
import dev.erst.gridgrind.core.validation.WorkbookStaticRequestContract;
import dev.erst.gridgrind.core.validation.WorkbookStaticViolation;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteCipher;
import dev.erst.gridgrind.core.vocabulary.ExcelOoxmlWriteHash;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.engine.backend.xssf.CellNotFoundException;
import dev.erst.gridgrind.engine.backend.xssf.InvalidCellAddressException;
import dev.erst.gridgrind.engine.backend.xssf.NamedRangeNotFoundException;
import dev.erst.gridgrind.engine.backend.xssf.SheetNotFoundException;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactIo;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import dev.erst.gridgrind.engine.formulas.MissingExternalWorkbookException;
import dev.erst.gridgrind.engine.formulas.UnregisteredUserDefinedFunctionException;
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
  void workbookCommandConverterRejectsIdentityMismatchesAndBroadNamedRangeDeletes() {
    IllegalArgumentException tableMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                DefaultGridGrindRequestExecutorTestSupport.command(
                    new TableSelector.ByNameOnSheet(
                        new TableName("OtherTable"), new SheetName("Budget")),
                    new StructuredMutationAction.SetTable(
                        TableSpec.withDefaultMetadata(
                            new TableName("BudgetTable"),
                            new SheetName("Budget"),
                            new RangeRef("A1:B2"),
                            false,
                            new TableStyleSpec.None()))));
    assertEquals(
        "SET_TABLE target must match table.name and table.sheetName", tableMismatch.getMessage());

    IllegalArgumentException tableSheetMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                DefaultGridGrindRequestExecutorTestSupport.command(
                    new TableSelector.ByNameOnSheet(
                        new TableName("BudgetTable"), new SheetName("Archive")),
                    new StructuredMutationAction.SetTable(
                        TableSpec.withDefaultMetadata(
                            new TableName("BudgetTable"),
                            new SheetName("Budget"),
                            new RangeRef("A1:B2"),
                            false,
                            new TableStyleSpec.None()))));
    assertEquals(
        "SET_TABLE target must match table.name and table.sheetName",
        tableSheetMismatch.getMessage());

    IllegalArgumentException pivotMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                DefaultGridGrindRequestExecutorTestSupport.command(
                    new PivotTableSelector.ByNameOnSheet("OtherPivot", new SheetName("Budget")),
                    new StructuredMutationAction.SetPivotTable(
                        new PivotTableSpec(
                            "SalesPivot",
                            new SheetName("Budget"),
                            new PivotTableSpec.PivotSourceSpec.NamedRange(
                                new DefinedName("BudgetSource")),
                            new PivotTableSpec.Anchor(new CellAddress("B3")),
                            List.of(),
                            List.of(),
                            List.of(),
                            List.of(
                                new PivotTableSpec.DataField(
                                    "Amount",
                                    dev.erst.gridgrind.core.vocabulary
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
                DefaultGridGrindRequestExecutorTestSupport.command(
                    new PivotTableSelector.ByNameOnSheet("SalesPivot", new SheetName("Archive")),
                    new StructuredMutationAction.SetPivotTable(
                        new PivotTableSpec(
                            "SalesPivot",
                            new SheetName("Budget"),
                            new PivotTableSpec.PivotSourceSpec.NamedRange(
                                new DefinedName("BudgetSource")),
                            new PivotTableSpec.Anchor(new CellAddress("B3")),
                            List.of(),
                            List.of(),
                            List.of(),
                            List.of(
                                new PivotTableSpec.DataField(
                                    "Amount",
                                    dev.erst.gridgrind.core.vocabulary
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
                DefaultGridGrindRequestExecutorTestSupport.command(
                    new NamedRangeSelector.SheetScope(
                        new DefinedName("BudgetTotal"), new SheetName("Archive")),
                    new StructuredMutationAction.SetNamedRange(
                        new DefinedName("BudgetTotal"),
                        new NamedRangeScope.Sheet(new SheetName("Budget")),
                        NamedRangeTarget.range(new SheetName("Budget"), new RangeRef("B4")))));
    assertEquals(
        "SET_NAMED_RANGE target must match action name and scope", namedRangeMismatch.getMessage());

    IllegalArgumentException namedRangeNameMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                DefaultGridGrindRequestExecutorTestSupport.command(
                    new NamedRangeSelector.SheetScope(
                        new DefinedName("OtherTotal"), new SheetName("Budget")),
                    new StructuredMutationAction.SetNamedRange(
                        new DefinedName("BudgetTotal"),
                        new NamedRangeScope.Sheet(new SheetName("Budget")),
                        NamedRangeTarget.range(new SheetName("Budget"), new RangeRef("B4")))));
    assertEquals(
        "SET_NAMED_RANGE target must match action name and scope",
        namedRangeNameMismatch.getMessage());

    IllegalArgumentException workbookScopeMismatch =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                DefaultGridGrindRequestExecutorTestSupport.command(
                    new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal")),
                    new StructuredMutationAction.SetNamedRange(
                        new DefinedName("BudgetTotal"),
                        new NamedRangeScope.Sheet(new SheetName("Budget")),
                        NamedRangeTarget.range(new SheetName("Budget"), new RangeRef("B4")))));
    assertEquals(
        "SET_NAMED_RANGE target must match action name and scope",
        workbookScopeMismatch.getMessage());

    IllegalArgumentException broadDelete =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                DefaultGridGrindRequestExecutorTestSupport.command(
                    new NamedRangeSelector.ByName(new DefinedName("BudgetTotal")),
                    new StructuredMutationAction.DeleteNamedRange()));
    assertEquals("Mutation target does not match DeleteNamedRange", broadDelete.getMessage());
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

    dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep enriched =
        assertInstanceOf(
            dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep.class,
            GridGrindProblems.enrichContext(
                new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep(
                    dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape
                        .known("NEW", "NONE"),
                    new dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces
                        .StepReference(0, new StepId("step"), "INSPECTION", "GET_NAMED_RANGES"),
                    dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.ProblemLocation
                        .unknown()),
                new NamedRangeNotFoundException(
                    "LocalTotal",
                    new NamedRangeScopeReport.Sheet(
                        new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget")))));
    assertEquals(
        java.util.Optional.of("Budget"),
        enriched.sheetName().map(dev.erst.gridgrind.core.vocabulary.WireString::value));
    assertEquals(
        java.util.Optional.of("A1"),
        assertInstanceOf(
                dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep.class,
                GridGrindProblems.enrichContext(
                    new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep(
                        dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape
                            .known("NEW", "NONE"),
                        new dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces
                            .StepReference(1, new StepId("cell"), "MUTATION", "SET_CELL"),
                        dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces
                            .ProblemLocation.unknown()),
                    new CellNotFoundException("A1")))
            .address());
    assertEquals(
        java.util.Optional.of("BAD!"),
        assertInstanceOf(
                dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep.class,
                GridGrindProblems.enrichContext(
                    new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteStep(
                        dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape
                            .known("NEW", "NONE"),
                        new dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces
                            .StepReference(2, new StepId("cell"), "MUTATION", "SET_CELL"),
                        dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces
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
    ExecutionModeSpec defaults = defaultModesRequest.effectiveExecutionMode();
    assertInstanceOf(ExecutionModeSpec.FullXssf.class, defaults);

    WorkbookPlan eventReadRequest =
        request(
            new WorkbookPlan.WorkbookSource.ExistingFile("/tmp/book.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionModeSpec.eventRead(),
            null,
            List.of(),
            List.of(
                inspect(
                    "workbook",
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.GetWorkbookSummary())));
    assertTrue(
        dev.erst.gridgrind.core.semantics.BackendCapabilities.eventRead()
            .accepts(eventReadRequest));

    WorkbookPlan streamingInspectionBeforeEnsure =
        request(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionModeSpec.streamingWrite(),
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
    WorkbookLifecycle.deleteIfExists(deleteTarget);
    assertFalse(Files.exists(deleteTarget));

    Path retained = Files.createTempFile("gridgrind-delete-retained-", ".tmp");
    WorkbookLifecycle.deleteIfExists(
        retained,
        ignored -> {
          throw new IOException("best effort");
        });
    assertTrue(Files.exists(retained));
    WorkbookLifecycle.deleteIfExists(retained);
    WorkbookLifecycle.deleteIfExists(
        null,
        ignored -> {
          throw new AssertionError("null paths must not invoke the deleter");
        });

    try (var prepared =
        ExecutionInputBindingsFixtureSupport.preparedBindings(
            workingDirectory,
            List.of(),
            java.util.Map.of(
                new dev.erst.gridgrind.core.source.SecretReference("output-password"),
                "output-secret"))) {
      ResolvedPersistenceSecurity noneOptions =
          ExecutionRequestPaths.persistenceOptions(
              new WorkbookPlan.WorkbookPersistence.None(), prepared.bindings());
      ResolvedPersistenceSecurity saveAsOptions =
          ExecutionRequestPaths.persistenceOptions(
              new WorkbookPlan.WorkbookPersistence.SaveAs(
                  "/tmp/out.xlsx",
                  WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                  new OoxmlPersistenceSecuritySpec(
                      new dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec
                          .Encrypt(
                          new OoxmlEncryptionSpec(
                              new dev.erst.gridgrind.core.source.SecretReference("output-password"),
                              ExcelOoxmlWriteCipher.AES_256,
                              ExcelOoxmlWriteHash.SHA_512)),
                      new dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec
                          .None())),
              prepared.bindings());
      ResolvedPersistenceSecurity overwriteOptions =
          ExecutionRequestPaths.persistenceOptions(
              new WorkbookPlan.WorkbookPersistence.Overwrite(
                  new OoxmlPersistenceSecuritySpec(
                      new dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec
                          .Encrypt(
                          new OoxmlEncryptionSpec(
                              new dev.erst.gridgrind.core.source.SecretReference("output-password"),
                              ExcelOoxmlWriteCipher.AES_256,
                              ExcelOoxmlWriteHash.SHA_512)),
                      new dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec
                          .None())),
              prepared.bindings());
      assertTrue(noneOptions.writesPlaintextUnsigned());
      assertFalse(saveAsOptions.writesPlaintextUnsigned());
      assertFalse(overwriteOptions.writesPlaintextUnsigned());
    }

    OoxmlPackageSecurityReport newSecurity =
        ExecutionRequestPaths.sourcePackageSecurity(new WorkbookPlan.WorkbookSource.New());
    OoxmlPackageSecurityReport existingSecurity =
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
    org.junit.jupiter.api.Assertions.assertThrows(
        IllegalStateException.class,
        () ->
            ExecutionRequestPaths.sourceEncryptionPassword(
                new WorkbookPlan.WorkbookSource.ExistingFile(
                    "/tmp/in.xlsx",
                    new OoxmlOpenSecuritySpec(
                        java.util.Optional.of(
                            new dev.erst.gridgrind.core.source.SecretReference("open-secret"))))));
  }

  private static void assertStreamingPersistenceBehaviors() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-streaming-policy-");
    WorkbookLifecycle workbookSupport =
        ExecutionContextFixtureSupport.workbookSupport(workingDirectory);
    Path materialized = createWorkbookFile("gridgrind-streaming-source-");
    try (var prepared = ExecutionInputBindingsFixtureSupport.preparedBindings(workingDirectory)) {
      WorkbookResultPersistence.PersistenceOutcome notSaved =
          workbookSupport.persistStreamingWorkbook(
              materialized,
              new WorkbookPlan.WorkbookPersistence.None(),
              new WorkbookPlan.WorkbookSource.New(),
              prepared.bindings(),
              StagedArtifactAcceptance.none());
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
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition.CREATE_NEW);
      WorkbookResultPersistence.PersistenceOutcome.SavedAs savedAs =
          assertInstanceOf(
              WorkbookResultPersistence.PersistenceOutcome.SavedAs.class,
              workbookSupport.persistStreamingWorkbook(
                  materialized,
                  new WorkbookPlan.WorkbookPersistence.SaveAs(
                      saveAsPath.toString(), WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
                  new WorkbookPlan.WorkbookSource.New(),
                  prepared.bindings(),
                  StagedArtifactAcceptance.none()));
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
              dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition
                  .REPLACE_EXISTING);
      WorkbookResultPersistence.PersistenceOutcome.Overwritten overwritten =
          assertInstanceOf(
              WorkbookResultPersistence.PersistenceOutcome.Overwritten.class,
              workbookSupport.persistStreamingWorkbook(
                  overwriteMaterialized,
                  new WorkbookPlan.WorkbookPersistence.Overwrite(
                      OoxmlPersistenceSecuritySpec.none()),
                  new WorkbookPlan.WorkbookSource.ExistingFile(overwriteSourcePath.toString()),
                  prepared.bindings(),
                  StagedArtifactAcceptance.none()));
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
                          OoxmlPersistenceSecuritySpec.none()),
                      new WorkbookPlan.WorkbookSource.New(),
                      prepared.bindings(),
                      StagedArtifactAcceptance.none()));
      assertEquals(
          "OVERWRITE persistence requires an EXISTING source", overwriteFailure.getMessage());
    }
  }

  private static void assertRuntimeGuardBehaviors() throws IOException {
    ExecutionResults responseSupport =
        new ExecutionResults(XssfWorkbook::close, WorkbookArtifactIo.MaterializedWorkbook::close);
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
                ProtocolVersion.CURRENT,
                request,
                ExecutionContextFixtureSupport.startJournal(request, ExecutionProgressSink.NOOP),
                () -> {
                  throw new UnsupportedOperationException("boom");
                }));
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, runtimeFailure.problem().code());

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      WorkbookResult.Failure workbookRuntimeFailure =
          assertInstanceOf(
              WorkbookResult.Failure.class,
              responseSupport.guardUnexpectedRuntime(
                  ProtocolVersion.CURRENT,
                  request,
                  ExecutionContextFixtureSupport.startJournal(request, ExecutionProgressSink.NOOP),
                  workbook,
                  () -> {
                    throw new UnsupportedOperationException("boom");
                  }));
      assertEquals(GridGrindProblemCode.INTERNAL_ERROR, workbookRuntimeFailure.problem().code());
    }

    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      ExecutionResults closeFailingResponseSupport =
          new ExecutionResults(
              ignored -> {
                throw new IOException("close failed");
              },
              WorkbookArtifactIo.MaterializedWorkbook::close);
      WorkbookResult.Failure closeFailure =
          assertInstanceOf(
              WorkbookResult.Failure.class,
              closeFailingResponseSupport.guardUnexpectedRuntime(
                  ProtocolVersion.CURRENT,
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
    CellMutationAction.SetCell setBlank = new CellMutationAction.SetCell(new CellValueSpec.Blank());
    CellMutationAction.SetCell setText = new CellMutationAction.SetCell(textCell("plain text"));
    CellMutationAction.SetCell setRichText =
        new CellMutationAction.SetCell(new CellValueSpec.RichText(List.of(richTextRun("rich"))));
    CellMutationAction.SetCell setNumeric =
        new CellMutationAction.SetCell(new CellValueSpec.NumberValue(12.5));
    CellMutationAction.SetCell setBoolean =
        new CellMutationAction.SetCell(new CellValueSpec.BooleanValue(true));
    CellMutationAction.SetCell setDate =
        new CellMutationAction.SetCell(new CellValueSpec.Date(LocalDate.parse("2026-04-17")));
    CellMutationAction.SetCell setDateTime =
        new CellMutationAction.SetCell(
            new CellValueSpec.DateTime(LocalDateTime.parse("2026-04-17T09:10:11")));
    CellMutationAction.SetCell setFormula =
        new CellMutationAction.SetCell(formulaCell("SUM(A1:A2)"));
    StructuredMutationAction.SetPivotTable pivotFromRange =
        pivotTableAction(
            new PivotTableSpec.PivotSourceSpec.Range(
                new SheetName("Budget"), new RangeRef("A1:B5")));
    StructuredMutationAction.SetPivotTable pivotFromNamedRange =
        pivotTableAction(
            new PivotTableSpec.PivotSourceSpec.NamedRange(new DefinedName("BudgetSource")));
    StructuredMutationAction.SetPivotTable pivotFromTable =
        pivotTableAction(new PivotTableSpec.PivotSourceSpec.Table(new TableName("BudgetTable")));
    StructuredMutationAction.SetNamedRange sheetScopedNamedRange =
        new StructuredMutationAction.SetNamedRange(
            new DefinedName("LocalTotal"),
            new NamedRangeScope.Sheet(new SheetName("Budget")),
            NamedRangeTarget.formula("SUM(Budget!B2:B4)"));
    StructuredMutationAction.SetNamedRange workbookScopedNamedRange =
        new StructuredMutationAction.SetNamedRange(
            new DefinedName("BudgetTotal"),
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
            new DrawingObjectSelector.AllOnSheet(new SheetName("Budget"))));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new DrawingObjectSelector.ByName(new SheetName("Budget"), "Logo")));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new ChartSelector.ByName(new SheetName("Budget"), "Revenue")));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new ChartSelector.AllOnSheet(new SheetName("Budget"))));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(new TableSelector.All()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new TableSelector.ByName(new TableName("BudgetTable"))));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new TableSelector.ByNames(
                List.of(new TableName("BudgetTable"), new TableName("ForecastTable")))));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new TableSelector.ByNameOnSheet(
                new TableName("BudgetTable"), new SheetName("Budget"))));
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
            new PivotTableSelector.ByNameOnSheet("SalesPivot", new SheetName("Budget"))));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(new NamedRangeSelector.All()));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new NamedRangeSelector.ByName(new DefinedName("BudgetTotal"))));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new NamedRangeSelector.ByNames(
                List.of(new DefinedName("BudgetTotal"), new DefinedName("ForecastTotal")))));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal"))));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new NamedRangeSelector.SheetScope(
                new DefinedName("BudgetTotal"), new SheetName("Budget"))));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new TableRowSelector.AllRows(
                new TableSelector.ByNameOnSheet(
                    new TableName("BudgetTable"), new SheetName("Budget")))));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new TableRowSelector.ByIndex(
                new TableSelector.ByNameOnSheet(
                    new TableName("BudgetTable"), new SheetName("Budget")),
                0)));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new TableCellSelector.ByColumnName(
                new TableRowSelector.ByKeyCell(
                    new TableSelector.ByNameOnSheet(
                        new TableName("BudgetTable"), new SheetName("Budget")),
                    "Item",
                    textCell("Hosting")),
                "Amount")));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new SheetSelector.ByNames(List.of(new SheetName("Budget")))));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(
            new SheetSelector.ByNames(
                List.of(new SheetName("Budget"), new SheetName("Forecast")))));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.sheetNameFor(selectorWithoutSheetOrAddress));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.addressFor(
            new SheetSelector.ByName(new SheetName("Budget"))));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.addressFor(selectorWithoutSheetOrAddress));

    assertEquals(
        java.util.Optional.of("B2"),
        ExecutionSelectorDiagnosticFields.addressFor(
            new RangeSelector.RectangularWindow(
                new SheetName("Budget"), new CellAddress("B2"), 2, 2)));
    assertEquals(
        java.util.Optional.of("A1:B2"),
        ExecutionSelectorDiagnosticFields.rangeFor(
            new RangeSelector.ByRanges(new SheetName("Budget"), List.of(new RangeRef("A1:B2")))));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionSelectorDiagnosticFields.rangeFor(
            new RangeSelector.ByRanges(
                new SheetName("Budget"), List.of(new RangeRef("A1:B2"), new RangeRef("C1:D2")))));
    assertEquals(
        java.util.Optional.of("B2:C3"),
        ExecutionSelectorDiagnosticFields.rangeFor(
            new RangeSelector.RectangularWindow(
                new SheetName("Budget"), new CellAddress("B2"), 2, 2)));
  }

  private static void assertNamedRangeSelectorDiagnostics() {
    assertEquals(
        Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector)
                new NamedRangeSelector.SheetScope(
                    new DefinedName("LocalTotal"), new SheetName("Budget"))));
    assertEquals(
        Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector)
                new NamedRangeSelector.AnyOf(
                    List.of(
                        new NamedRangeSelector.SheetScope(
                            new DefinedName("LocalTotal"), new SheetName("Budget"))))));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.SheetScope(
                        new DefinedName("LocalTotal"), new SheetName("Budget")),
                    new NamedRangeSelector.SheetScope(
                        new DefinedName("ForecastTotal"), new SheetName("Forecast"))))));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector) new NamedRangeSelector.All()));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector) new NamedRangeSelector.ByName(new DefinedName("BudgetTotal"))));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector)
                new NamedRangeSelector.ByNames(List.of(new DefinedName("BudgetTotal")))));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector)
                new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal"))));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector.Ref)
                new NamedRangeSelector.ByName(new DefinedName("BudgetTotal"))));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector.Ref)
                new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal"))));
    assertEquals(
        Optional.of("Budget"),
        ExecutionSelectorDiagnosticFields.singleSheetName(
            (NamedRangeSelector.Ref)
                new NamedRangeSelector.SheetScope(
                    new DefinedName("BudgetTotal"), new SheetName("Budget"))));

    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(new NamedRangeSelector.All()));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector) new NamedRangeSelector.ByName(new DefinedName("BudgetTotal"))));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            new NamedRangeSelector.ByNames(List.of(new DefinedName("BudgetTotal")))));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            new NamedRangeSelector.ByNames(
                List.of(new DefinedName("BudgetTotal"), new DefinedName("ForecastTotal")))));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector)
                new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal"))));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector)
                new NamedRangeSelector.SheetScope(
                    new DefinedName("BudgetTotal"), new SheetName("Budget"))));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            new NamedRangeSelector.AnyOf(
                List.of(new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal"))))));
    assertEquals(
        Optional.empty(),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal")),
                    new NamedRangeSelector.WorkbookScope(new DefinedName("ForecastTotal"))))));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector.Ref)
                new NamedRangeSelector.ByName(new DefinedName("BudgetTotal"))));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector.Ref)
                new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal"))));
    assertEquals(
        Optional.of("BudgetTotal"),
        ExecutionSelectorDiagnosticFields.singleNamedRangeName(
            (NamedRangeSelector.Ref)
                new NamedRangeSelector.SheetScope(
                    new DefinedName("BudgetTotal"), new SheetName("Budget"))));
  }

  private static void assertExceptionAndStepDiagnostics() {
    CellMutationAction.SetCell setText = new CellMutationAction.SetCell(textCell("plain text"));
    StructuredMutationAction.SetPivotTable pivotFromRange =
        pivotTableAction(
            new PivotTableSpec.PivotSourceSpec.Range(
                new SheetName("Budget"), new RangeRef("A1:B5")));
    MutationStep pivotStep =
        new MutationStep(
            new StepId("pivot-step"),
            new PivotTableSelector.ByNameOnSheet("SalesPivot", new SheetName("Budget")),
            pivotFromRange);
    MutationStep addressedCellStep =
        new MutationStep(
            new StepId("cell-step"),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("D4")),
            setText);

    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionExceptionDiagnosticFields.sheetNameFor(new SheetNotFoundException("Budget")));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionExceptionDiagnosticFields.sheetNameFor(new RuntimeException("x")));
    assertEquals(
        java.util.Optional.empty(),
        ExecutionExceptionDiagnosticFields.sheetNameFor(
            new NamedRangeNotFoundException("BudgetTotal", new NamedRangeScopeReport.Workbook())));
    assertEquals(
        java.util.Optional.of("Budget"),
        ExecutionExceptionDiagnosticFields.sheetNameFor(
            new NamedRangeNotFoundException(
                "LocalTotal",
                new NamedRangeScopeReport.Sheet(
                    new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget")))));
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
      PivotTableSpec.PivotSourceSpec source) {
    return new StructuredMutationAction.SetPivotTable(
        new PivotTableSpec(
            "SalesPivot",
            new SheetName("Budget"),
            source,
            new PivotTableSpec.Anchor(new CellAddress("C5")),
            List.of(),
            List.of(),
            List.of(),
            List.of(
                new PivotTableSpec.DataField(
                    "Amount",
                    dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction.SUM,
                    "Amount",
                    Optional.empty()))));
  }

  private static StructuredMutationAction.SetConditionalFormatting conditionalFormattingAction() {
    return new StructuredMutationAction.SetConditionalFormatting(
        new ConditionalFormattingSpec(
            List.of(
                new ConditionalFormattingRuleSpec.FormulaRule(
                    "B2>0",
                    true,
                    java.util.Optional.of(
                        new DifferentialStylePatch(
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
