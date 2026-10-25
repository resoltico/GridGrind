package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.CalculationPolicyInput;
import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentInput;
import dev.erst.gridgrind.core.plan.ExecutionJournalInput;
import dev.erst.gridgrind.core.plan.ExecutionModeInput;
import dev.erst.gridgrind.core.plan.ExecutionPolicyInput;
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
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;
import java.util.stream.Stream;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Proves a private-source race remains a structured open-workbook failure. */
class DefaultGridGrindRequestExecutorPrivateSourceRaceTest {
  @TempDir Path root;

  @Test
  void reportsAnOpenWorkbookProblemWhenTheVerifiedPrivateSourceDisappears() throws Exception {
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
            new ExecutionPolicyInput(
                ExecutionModeInput.defaults(),
                new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE),
                CalculationPolicyInput.defaults(),
                dev.erst.gridgrind.core.plan.AssertionModeInput.defaults()),
            FormulaEnvironmentInput.empty(),
            List.of());
    AtomicBoolean removedPrivateSource = new AtomicBoolean();

    WorkbookResult response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request,
                new ExecutionInputBindings(
                    root,
                    tempRoot,
                    ExecutionGrantTestSupport.noPublication(
                        List.of(), List.of(root.resolve("source.xlsx")))),
                event ->
                    removeMaterializedSourceAfterPreflight(event, tempRoot, removedPrivateSource));

    WorkbookResult.Failure failure = assertInstanceOf(WorkbookResult.Failure.class, response);
    assertTrue(removedPrivateSource.get());
    assertEquals(GridGrindProblemCode.WORKBOOK_NOT_FOUND, failure.problem().code());
    assertInstanceOf(ProblemContext.OpenWorkbook.class, failure.problem().context());
  }

  @Test
  void failsClosedWhenTheDirectReadMaterializationDisappearsAfterPreflight() throws Exception {
    createWorkbook(root.resolve("source.xlsx"));
    Path tempRoot = Files.createDirectory(root.resolve("private-temp"));
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("source.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None(),
            verbosePolicy(ExecutionModeInput.eventRead()),
            FormulaEnvironmentInput.empty(),
            List.of(
                new InspectionStep(
                    "summary",
                    new WorkbookSelector.Current(),
                    new WorkbookIntrospectionQuery.GetWorkbookSummary())));
    AtomicBoolean removedPrivateSource = new AtomicBoolean();

    WorkbookResult response =
        new DefaultGridGrindRequestExecutor()
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
    assertEquals(GridGrindProblemCode.WORKBOOK_NOT_FOUND, failure.problem().code());
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
            verbosePolicy(ExecutionModeInput.streamingWrite()),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "ensure-sheet",
                    new SheetSelector.ByName("Budget"),
                    new WorkbookMutationAction.EnsureSheet())));
    AtomicBoolean replacedOutputDirectory = new AtomicBoolean();

    WorkbookResult response =
        new DefaultGridGrindRequestExecutor()
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
    assertEquals(GridGrindProblemCode.PUBLICATION_UNCERTAIN, failure.problem().code());
    assertInstanceOf(ProblemContext.PersistWorkbook.class, failure.problem().context());
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
            verbosePolicy(ExecutionModeInput.defaults()),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "ensure-sheet",
                    new SheetSelector.ByName("Budget"),
                    new WorkbookMutationAction.EnsureSheet())));
    AtomicBoolean replacedOutputDirectory = new AtomicBoolean();

    WorkbookResult response =
        new DefaultGridGrindRequestExecutor()
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
    assertEquals(GridGrindProblemCode.PUBLICATION_UNCERTAIN, failure.problem().code());
    assertInstanceOf(ProblemContext.PersistWorkbook.class, failure.problem().context());
  }

  private static ExecutionPolicyInput verbosePolicy(ExecutionModeInput mode) {
    return new ExecutionPolicyInput(
        mode,
        new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE),
        CalculationPolicyInput.defaults(),
        dev.erst.gridgrind.core.plan.AssertionModeInput.defaults());
  }

  private static void createWorkbook(Path source) throws java.io.IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        var output = Files.newOutputStream(source)) {
      workbook.createSheet("Budget");
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
    try (Stream<Path> files = Files.list(tempRoot)) {
      files
          .filter(path -> path.getFileName().toString().startsWith("gridgrind-source-workbook-"))
          .findFirst()
          .ifPresent(
              source -> {
                try {
                  Files.delete(source);
                  removedPrivateSource.set(true);
                } catch (java.io.IOException exception) {
                  throw new IllegalStateException(
                      "test could not remove private source", exception);
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
      replacedOutputDirectory.set(true);
    } catch (java.io.IOException exception) {
      throw new IllegalStateException("test could not replace bound output directory", exception);
    }
  }
}
