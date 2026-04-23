package dev.erst.gridgrind.authoring;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.ExecutionJournalInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput.ReadMode;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput.WriteMode;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.executor.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.executor.ExecutionInputBindings;
import dev.erst.gridgrind.executor.ExecutionJournalSink;
import java.io.ByteArrayOutputStream;
import java.lang.reflect.Method;
import java.net.URL;
import java.net.URLClassLoader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Locale;
import java.util.stream.Collectors;
import javax.tools.DiagnosticCollector;
import javax.tools.JavaCompiler;
import javax.tools.JavaFileObject;
import javax.tools.StandardJavaFileManager;
import javax.tools.ToolProvider;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Focused end-to-end tests for the Java authoring API boundary and example workflow. */
class GridGrindPlanTest {
  @TempDir Path tempDir;

  @Test
  void serializesOneFluentPlanBackToTheCanonicalWorkbookPlan() throws Exception {
    assertEquals(
        "planId must not be blank",
        assertThrows(IllegalArgumentException.class, () -> GridGrindPlan.newWorkbook().planId(" "))
            .getMessage());

    GridGrindPlan plan =
        GridGrindPlan.newWorkbook()
            .planId("budget-plan")
            .saveAs(tempDir.resolve("budget.xlsx"))
            .journal(ExecutionJournalLevel.VERBOSE)
            .mutate(Targets.sheet("Budget").ensureExists())
            .mutate(Targets.cell("Budget", "A1").set(Values.text("Owner")))
            .inspect(Targets.cell("Budget", "A1").read())
            .assertThat(Targets.cell("Budget", "A1").valueEquals(Values.expectedText("Owner")));

    WorkbookPlan canonical = plan.toPlan();
    assertEquals("budget-plan", canonical.planId());
    assertEquals("mutation-001", canonical.steps().get(0).stepId());
    assertEquals("mutation-002", canonical.steps().get(1).stepId());
    assertEquals("inspection-001", canonical.steps().get(2).stepId());
    assertEquals("assertion-001", canonical.steps().get(3).stepId());

    WorkbookPlan reread = GridGrindJson.readRequest(plan.toJsonBytes());
    assertEquals(canonical, reread);
    assertTrue(plan.toJsonString().contains("\"planId\" : \"budget-plan\""));
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    plan.writeJson(outputStream);
    assertTrue(
        outputStream.toString(StandardCharsets.UTF_8).contains("\"planId\" : \"budget-plan\""));
  }

  @Test
  void preservesImportedCanonicalExecutionFieldsAtTheBoundary() {
    WorkbookPlan canonical =
        new WorkbookPlan(
            dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
            "canonical-plan",
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            new ExecutionPolicyInput(
                null,
                new dev.erst.gridgrind.contract.dto.ExecutionJournalInput(
                    ExecutionJournalLevel.NORMAL),
                null),
            new FormulaEnvironmentInput(List.of(), null, List.of()),
            List.of());

    WorkbookPlan roundTrip = GridGrindPlan.from(canonical).toPlan();
    assertEquals(canonical, roundTrip);
  }

  @Test
  void journalPreservesImportedExecutionModeAndCalculation() {
    ExecutionModeInput mode = new ExecutionModeInput(ReadMode.EVENT_READ, WriteMode.FULL_XSSF);
    dev.erst.gridgrind.contract.dto.CalculationPolicyInput calculation =
        new dev.erst.gridgrind.contract.dto.CalculationPolicyInput(
            new dev.erst.gridgrind.contract.dto.CalculationStrategyInput.DoNotCalculate(), true);
    WorkbookPlan canonical =
        new WorkbookPlan(
            dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
            "canonical-plan",
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            new ExecutionPolicyInput(
                mode, new ExecutionJournalInput(ExecutionJournalLevel.NORMAL), calculation),
            null,
            List.of());

    WorkbookPlan journaled =
        GridGrindPlan.from(canonical).journal(ExecutionJournalLevel.VERBOSE).toPlan();
    assertEquals(mode, journaled.execution().mode());
    assertEquals(calculation, journaled.execution().calculation());
    assertEquals(ExecutionJournalLevel.VERBOSE, journaled.execution().journal().level());
  }

  @Test
  void importedPlansContinueAutoGeneratedCountersAndRespectNamedSteps() {
    WorkbookPlan imported =
        new WorkbookPlan(
            dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
            "imported",
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            null,
            null,
            List.of(
                new MutationStep(
                    "mutation-001",
                    Targets.sheet("Budget").selector(),
                    new dev.erst.gridgrind.contract.action.MutationAction.EnsureSheet()),
                new InspectionStep(
                    "inspection-001",
                    Targets.cell("Budget", "A1").selector(),
                    new dev.erst.gridgrind.contract.query.InspectionQuery.GetCells()),
                new AssertionStep(
                    "assertion-001",
                    Targets.cell("Budget", "A1").selector(),
                    new dev.erst.gridgrind.contract.assertion.Assertion.CellValue(
                        Values.toExpectedCellValue(Values.expectedBlank())))));

    WorkbookPlan plan =
        GridGrindPlan.from(imported)
            .planId(null)
            .mutate(Targets.sheet("Budget").renameTo("Budget 2026"))
            .inspect(Targets.cell("Budget 2026", "A1").read())
            .assertThat(Targets.cell("Budget 2026", "A1").valueEquals(Values.expectedBlank()))
            .mutate(Targets.sheet("Budget 2026").ensureExists().named("keep-sheet"))
            .toPlan();

    assertEquals("mutation-002", plan.steps().get(3).stepId());
    assertEquals("inspection-002", plan.steps().get(4).stepId());
    assertEquals("assertion-002", plan.steps().get(5).stepId());
    assertEquals("keep-sheet", plan.steps().get(6).stepId());
  }

  @Test
  void conveniencePersistenceMethodsCoverInMemoryAndOverwriteBranches() {
    GridGrindPlan plan = GridGrindPlan.newWorkbook();
    assertEquals(plan, plan.planId(null));
    assertInstanceOf(
        WorkbookPlan.WorkbookPersistence.None.class,
        plan.saveAs(tempDir.resolve("copy.xlsx")).inMemoryOnly().toPlan().persistence());
    assertInstanceOf(
        WorkbookPlan.WorkbookPersistence.OverwriteSource.class,
        GridGrindPlan.open(tempDir.resolve("source.xlsx"))
            .overwriteSource()
            .toPlan()
            .persistence());
  }

  @Test
  void runsOneTableAwareOfficeWorkflowWithExplicitExecutor() throws Exception {
    Files.writeString(tempDir.resolve("item.txt"), "Hosting");
    Path outputPath = tempDir.resolve("budget.xlsx");

    GridGrindPlan plan =
        GridGrindPlan.newWorkbook()
            .saveAs(outputPath)
            .journal(ExecutionJournalLevel.VERBOSE)
            .mutate(Targets.sheet("Budget").ensureExists())
            .mutate(
                Targets.range("Budget", "A1:B3")
                    .setRows(
                        List.of(
                            Values.row(Values.text("Item"), Values.text("Amount")),
                            Values.row(Values.text("Hosting"), Values.number(100.0)),
                            Values.row(Values.text("Travel"), Values.number(50.0)))))
            .mutate(
                Targets.tableOnSheet("BudgetTable", "Budget")
                    .define(
                        Tables.define("BudgetTable", "Budget", "A1:B3", false, Tables.noStyle())))
            .mutate(
                Targets.table("BudgetTable")
                    .rowByKey("Item", Values.textFile(Path.of("item.txt")))
                    .cell("Amount")
                    .set(Values.number(125.0)))
            .inspect(
                Targets.table("BudgetTable")
                    .rowByKey("Item", Values.textFile(Path.of("item.txt")))
                    .cell("Amount")
                    .read())
            .assertThat(
                Targets.table("BudgetTable")
                    .rowByKey("Item", Values.textFile(Path.of("item.txt")))
                    .cell("Amount")
                    .valueEquals(Values.expectedNumber(125.0)));

    GridGrindResponse.Success response =
        assertInstanceOf(
            GridGrindResponse.Success.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    plan.toPlan(),
                    new ExecutionInputBindings(tempDir, (byte[]) null),
                    ExecutionJournalSink.NOOP));

    assertTrue(Files.exists(outputPath));
    assertEquals(1, response.assertions().size());
    assertNotNull(response.journal());
    InspectionResult.CellsResult cellsResult =
        assertInstanceOf(InspectionResult.CellsResult.class, response.inspections().getFirst());
    GridGrindResponse.CellReport.NumberReport numberReport =
        assertInstanceOf(
            GridGrindResponse.CellReport.NumberReport.class, cellsResult.cells().getFirst());
    assertEquals("Budget", cellsResult.sheetName());
    assertEquals("B2", numberReport.address());
    assertEquals(125.0, numberReport.numberValue());
  }

  @Test
  void shippedJavaAuthoringExampleCompilesAgainstThePublishedApi() throws Exception {
    JavaCompiler compiler = ToolProvider.getSystemJavaCompiler();
    assertNotNull(compiler, "system Java compiler must be available for example compilation");

    Path repoRoot = locateRepoRoot();
    Path examplePath = repoRoot.resolve("examples").resolve("java-authoring-workflow.java");
    assertTrue(Files.exists(examplePath), "examples/java-authoring-workflow.java must exist");
    Path bridgeSourcePath = tempDir.resolve("JavaAuthoringWorkflowExampleBridge.java");
    Files.writeString(
        bridgeSourcePath,
        """
        import dev.erst.gridgrind.contract.dto.GridGrindResponse;
        import java.nio.file.Path;

        public final class JavaAuthoringWorkflowExampleBridge {
          private JavaAuthoringWorkflowExampleBridge() {}

          public static GridGrindResponse run(Path workspace) throws Exception {
            return JavaAuthoringWorkflowExample.run(workspace);
          }
        }
        """);
    Path classesDirectory = Files.createDirectories(tempDir.resolve("example-classes"));
    DiagnosticCollector<JavaFileObject> diagnostics = new DiagnosticCollector<>();

    try (StandardJavaFileManager fileManager =
        compiler.getStandardFileManager(null, Locale.ROOT, StandardCharsets.UTF_8)) {
      Iterable<? extends JavaFileObject> compilationUnits =
          fileManager.getJavaFileObjects(examplePath.toFile(), bridgeSourcePath.toFile());
      boolean compiled =
          compiler
              .getTask(
                  null,
                  fileManager,
                  diagnostics,
                  List.of(
                      "--release",
                      "26",
                      "-proc:none",
                      "-classpath",
                      System.getProperty("java.class.path"),
                      "-d",
                      classesDirectory.toString()),
                  null,
                  compilationUnits)
              .call();
      assertTrue(
          compiled,
          () ->
              "examples/java-authoring-workflow.java must compile: "
                  + diagnostics.getDiagnostics().stream()
                      .map(
                          diagnostic ->
                              diagnostic.getKind()
                                  + " line "
                                  + diagnostic.getLineNumber()
                                  + ": "
                                  + diagnostic.getMessage(Locale.ROOT))
                      .collect(Collectors.joining(" | ")));
    }

    Path authoredInputs = Files.createDirectories(tempDir.resolve("authored-inputs"));
    Files.writeString(authoredInputs.resolve("item.txt"), "Hosting");

    ClassLoader parentClassLoader = Thread.currentThread().getContextClassLoader();
    try (URLClassLoader classLoader =
        new URLClassLoader(new URL[] {classesDirectory.toUri().toURL()}, parentClassLoader)) {
      Class<?> exampleClass = classLoader.loadClass("JavaAuthoringWorkflowExampleBridge");
      Method runMethod = exampleClass.getMethod("run", Path.class);
      GridGrindResponse.Success response =
          assertInstanceOf(GridGrindResponse.Success.class, runMethod.invoke(null, tempDir));
      assertTrue(Files.exists(tempDir.resolve("budget.xlsx")));
      assertEquals(1, response.assertions().size());
      assertEquals(1, response.inspections().size());
    }
  }

  private static Path locateRepoRoot() {
    Path current = Path.of("").toAbsolutePath();
    while (current != null && !Files.exists(current.resolve("settings.gradle.kts"))) {
      current = current.getParent();
    }
    assertNotNull(current, "test must run inside the GridGrind repository");
    return current;
  }
}
