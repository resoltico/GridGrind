package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournalInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
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
                CalculationPolicyInput.defaults()),
            FormulaEnvironmentInput.empty(),
            List.of());
    AtomicBoolean removedPrivateSource = new AtomicBoolean();

    WorkbookResult response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request,
                new ExecutionInputBindings(root, tempRoot),
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
                new ExecutionInputBindings(root, tempRoot),
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
                new ExecutionInputBindings(root, tempRoot),
                event ->
                    replaceBoundOutputDirectory(event, outputDirectory, replacedOutputDirectory));

    WorkbookResult.Failure failure = assertInstanceOf(WorkbookResult.Failure.class, response);
    assertTrue(replacedOutputDirectory.get());
    assertEquals(GridGrindProblemCode.UNSAFE_PATH_ACCESS, failure.problem().code());
    assertInstanceOf(ProblemContext.PersistWorkbook.class, failure.problem().context());
  }

  private static ExecutionPolicyInput verbosePolicy(ExecutionModeInput mode) {
    return new ExecutionPolicyInput(
        mode,
        new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE),
        CalculationPolicyInput.defaults());
  }

  private static void createWorkbook(Path source) throws java.io.IOException {
    try (XSSFWorkbook workbook = new XSSFWorkbook();
        var output = Files.newOutputStream(source)) {
      workbook.createSheet("Budget");
      workbook.write(output);
    }
  }

  private static void removeMaterializedSourceAfterPreflight(
      dev.erst.gridgrind.contract.dto.ExecutionJournal.Event event,
      Path tempRoot,
      AtomicBoolean removedPrivateSource) {
    if (!"RESOLVE_INPUTS".equals(event.category()) || !"succeeded".equals(event.detail())) {
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
      dev.erst.gridgrind.contract.dto.ExecutionJournal.Event event,
      Path outputDirectory,
      AtomicBoolean replacedOutputDirectory) {
    if (!"RESOLVE_INPUTS".equals(event.category()) || !"succeeded".equals(event.detail())) {
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
