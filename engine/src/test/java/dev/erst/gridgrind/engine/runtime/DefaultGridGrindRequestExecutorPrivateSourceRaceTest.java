package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.CalculationPolicySpec;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentSpec;
import dev.erst.gridgrind.core.plan.ExecutionJournalSpec;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.ExecutionPolicySpec;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.stream.Stream;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies retained-source semantics and publication rejection after actual topology changes. */
class DefaultGridGrindRequestExecutorPrivateSourceRaceTest {
  @TempDir Path root;

  @Test
  void retainedFullWorkbookUsesAdmittedBytesAfterPrivateUnlinkAndOriginalReplacement()
      throws Exception {
    Path source = root.resolve("source.xlsx");
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        var output = Files.newOutputStream(source)) {
      workbook.createSheet("Budget");
      workbook.write(output);
    }
    Path tempRoot = Files.createDirectory(root.resolve("private-temp"));
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("source.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None(),
            new ExecutionPolicySpec(
                ExecutionModeSpec.defaults(),
                new ExecutionJournalSpec(ExecutionJournalLevel.VERBOSE),
                CalculationPolicySpec.defaults(),
                dev.erst.gridgrind.core.plan.AssertionModeSpec.defaults()),
            FormulaEnvironmentSpec.empty(),
            List.of(
                new InspectionStep(
                    new StepId("summary"),
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.GetWorkbookSummary())));
    AtomicBoolean removedPrivateSource = new AtomicBoolean();

    WorkbookResult response =
        new DefaultGridGrindRequestExecutor(
                NativeInvocationContext.enterOwned(
                    dev.erst.gridgrind.core.semantics.ExecutionConventions.STANDARD))
            .execute(
                request,
                new ExecutionInputBindings(
                    root,
                    tempRoot,
                    ExecutionGrantTestSupport.noPublication(
                        List.of("GET_WORKBOOK_SUMMARY"), List.of(root.resolve("source.xlsx")))),
                event -> {
                  removeMaterializedSourceAfterPreflight(event, tempRoot, removedPrivateSource);
                  if (event.category()
                          == dev.erst.gridgrind.core.result.ExecutionProgressEvent.Category
                              .RESOLVE_INPUTS
                      && event.status()
                          == dev.erst.gridgrind.core.result.ExecutionProgressEvent.Status
                              .SUCCEEDED) {
                    try {
                      createWorkbook(source, "Changed");
                    } catch (java.io.IOException failure) {
                      throw new AssertionError(failure);
                    }
                  }
                });

    WorkbookResult.Success success = assertInstanceOf(WorkbookResult.Success.class, response);
    assertTrue(removedPrivateSource.get());
    var summary =
        assertInstanceOf(
                dev.erst.gridgrind.core.plan.WorkbookInspectionResult.WorkbookSummaryResult.class,
                success.inspections().getFirst())
            .workbook();
    assertEquals(
        List.of(new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget")), summary.sheetNames());
    try (var changed = new XSSFWorkbook(Files.newInputStream(source))) {
      assertEquals("Changed", changed.getSheetName(0));
    }
  }

  @Test
  void failsClosedWhenTheDirectReadMaterializationDisappearsAfterPreflight() throws Exception {
    createWorkbook(root.resolve("source.xlsx"), "Budget");
    Path tempRoot = Files.createDirectory(root.resolve("private-temp"));
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("source.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None(),
            verbosePolicy(ExecutionModeSpec.eventRead()),
            FormulaEnvironmentSpec.empty(),
            List.of(
                new InspectionStep(
                    new StepId("summary"),
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.GetWorkbookSummary())));
    AtomicBoolean removedPrivateSource = new AtomicBoolean();

    WorkbookResult response =
        new DefaultGridGrindRequestExecutor(
                NativeInvocationContext.enterOwned(
                    dev.erst.gridgrind.core.semantics.ExecutionConventions.STANDARD))
            .execute(
                request,
                new ExecutionInputBindings(
                    root,
                    tempRoot,
                    ExecutionGrantTestSupport.noPublication(
                        List.of("GET_WORKBOOK_SUMMARY"), List.of(root.resolve("source.xlsx")))),
                event ->
                    removeMaterializedSourceAfterPreflight(event, tempRoot, removedPrivateSource));

    WorkbookResult.Failure failure = assertInstanceOf(WorkbookResult.Failure.class, response);
    assertTrue(removedPrivateSource.get());
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertInstanceOf(ProblemContext.OpenWorkbook.class, failure.problem().context());
  }

  @Test
  void failsClosedWhenAStreamingOutputParentChangesAfterPreflight() throws Exception {
    Path outputDirectory = Files.createDirectory(root.resolve("output"));
    Path tempRoot = Files.createDirectory(root.resolve("private-temp"));
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "output/result.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            verbosePolicy(ExecutionModeSpec.streamingWrite()),
            FormulaEnvironmentSpec.empty(),
            List.of(
                new MutationStep(
                    new StepId("ensure-sheet"),
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new WorkbookMutationAction.EnsureSheet())));
    AtomicBoolean replacedOutputDirectory = new AtomicBoolean();

    WorkbookResult response =
        new DefaultGridGrindRequestExecutor(
                NativeInvocationContext.enterOwned(
                    dev.erst.gridgrind.core.semantics.ExecutionConventions.STANDARD))
            .execute(
                request,
                new ExecutionInputBindings(
                    root,
                    tempRoot,
                    ExecutionGrantTestSupport.saveAs(
                        List.of("ENSURE_SHEET"),
                        List.of(),
                        outputDirectory.resolve("result.xlsx"),
                        WorkbookPlan.WorkbookPersistence.IfExists.REJECT)),
                event ->
                    replaceBoundOutputDirectory(event, outputDirectory, replacedOutputDirectory));

    WorkbookResult.Failure failure = assertInstanceOf(WorkbookResult.Failure.class, response);
    assertTrue(replacedOutputDirectory.get());
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertInstanceOf(ProblemContext.PersistWorkbook.class, failure.problem().context());
    assertInstanceOf(
        dev.erst.gridgrind.core.result.WorkbookResultPersistence.PublicationOutcome.NotAttempted
            .class,
        assertInstanceOf(
                dev.erst.gridgrind.core.result.WorkbookResultPersistence.PersistenceOutcome.SavedAs
                    .class,
                failure.persistence())
            .publication());
    assertTrue(Files.notExists(outputDirectory.resolve("result.xlsx")));
    assertTrue(Files.notExists(root.resolve("output-replaced/result.xlsx")));
    assertEquals("competitor", Files.readString(outputDirectory.resolve("marker.txt")));
  }

  @Test
  void failsClosedWhenAFullXssfOutputParentChangesAfterPreflight() throws Exception {
    Path outputDirectory = Files.createDirectory(root.resolve("output"));
    Path tempRoot = Files.createDirectory(root.resolve("private-temp"));
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "output/result.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
            verbosePolicy(ExecutionModeSpec.defaults()),
            FormulaEnvironmentSpec.empty(),
            List.of(
                new MutationStep(
                    new StepId("ensure-sheet"),
                    new SheetSelector.ByName(new SheetName("Budget")),
                    new WorkbookMutationAction.EnsureSheet())));
    AtomicBoolean replacedOutputDirectory = new AtomicBoolean();

    WorkbookResult response =
        new DefaultGridGrindRequestExecutor(
                NativeInvocationContext.enterOwned(
                    dev.erst.gridgrind.core.semantics.ExecutionConventions.STANDARD))
            .execute(
                request,
                new ExecutionInputBindings(
                    root,
                    tempRoot,
                    ExecutionGrantTestSupport.saveAs(
                        List.of("ENSURE_SHEET"),
                        List.of(),
                        outputDirectory.resolve("result.xlsx"),
                        WorkbookPlan.WorkbookPersistence.IfExists.REJECT)),
                event ->
                    replaceBoundOutputDirectory(event, outputDirectory, replacedOutputDirectory));

    WorkbookResult.Failure failure = assertInstanceOf(WorkbookResult.Failure.class, response);
    assertTrue(replacedOutputDirectory.get());
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertInstanceOf(ProblemContext.PersistWorkbook.class, failure.problem().context());
    assertInstanceOf(
        dev.erst.gridgrind.core.result.WorkbookResultPersistence.PublicationOutcome.NotAttempted
            .class,
        assertInstanceOf(
                dev.erst.gridgrind.core.result.WorkbookResultPersistence.PersistenceOutcome.SavedAs
                    .class,
                failure.persistence())
            .publication());
    assertTrue(Files.notExists(outputDirectory.resolve("result.xlsx")));
    assertTrue(Files.notExists(root.resolve("output-replaced/result.xlsx")));
    assertEquals("competitor", Files.readString(outputDirectory.resolve("marker.txt")));
  }

  private static ExecutionPolicySpec verbosePolicy(ExecutionModeSpec mode) {
    return new ExecutionPolicySpec(
        mode,
        new ExecutionJournalSpec(ExecutionJournalLevel.VERBOSE),
        CalculationPolicySpec.defaults(),
        dev.erst.gridgrind.core.plan.AssertionModeSpec.defaults());
  }

  private static void createWorkbook(Path source, String sheetName) throws java.io.IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        var output = Files.newOutputStream(source)) {
      workbook.createSheet(sheetName);
      workbook.write(output);
    }
  }

  private static void removeMaterializedSourceAfterPreflight(
      dev.erst.gridgrind.core.result.ExecutionProgressEvent event,
      Path tempRoot,
      AtomicBoolean removedPrivateSource) {
    if (event.category()
            != dev.erst.gridgrind.core.result.ExecutionProgressEvent.Category.RESOLVE_INPUTS
        || event.status()
            != dev.erst.gridgrind.core.result.ExecutionProgressEvent.Status.SUCCEEDED) {
      return;
    }
    try (Stream<Path> files = Files.walk(tempRoot)) {
      files
          .filter(path -> path.getFileName().toString().startsWith("gridgrind-source-workbook-"))
          .findFirst()
          .ifPresent(source -> {
            try {
              Files.delete(source);
              removedPrivateSource.set(true);
            } catch (java.io.IOException exception) {
              throw new IllegalStateException("test could not remove private source", exception);
            }
          });
    } catch (java.io.IOException exception) {
      throw new IllegalStateException("test could not inspect private source", exception);
    }
  }

  private static void replaceBoundOutputDirectory(
      dev.erst.gridgrind.core.result.ExecutionProgressEvent event,
      Path outputDirectory,
      AtomicBoolean replacedOutputDirectory) {
    if (event.category()
            != dev.erst.gridgrind.core.result.ExecutionProgressEvent.Category.RESOLVE_INPUTS
        || event.status()
            != dev.erst.gridgrind.core.result.ExecutionProgressEvent.Status.SUCCEEDED) {
      return;
    }
    try {
      Files.move(outputDirectory, outputDirectory.resolveSibling("output-replaced"));
      Files.createDirectory(outputDirectory);
      Files.writeString(outputDirectory.resolve("marker.txt"), "competitor");
      replacedOutputDirectory.set(true);
    } catch (java.io.IOException exception) {
      throw new IllegalStateException("test could not replace bound output directory", exception);
    }
  }
}
