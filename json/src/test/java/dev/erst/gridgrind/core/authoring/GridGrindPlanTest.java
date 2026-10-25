package dev.erst.gridgrind.core.authoring;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.FormulaEnvironmentInput;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.ExecutionPolicyInput;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.SheetInspectionResult;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.api.GridGrindEngine;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import dev.erst.gridgrind.engine.api.GridGrindProgressSink;
import dev.erst.gridgrind.engine.api.GridGrindRequestInputs;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import java.io.ByteArrayOutputStream;
import java.lang.reflect.Method;
import java.net.URL;
import java.net.URLClassLoader;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Locale;
import java.util.Optional;
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
        assertThrows(
                IllegalArgumentException.class,
                () -> GridGrindPlan.newWorkbook().planId(new PlanId(" ")))
            .getMessage());

    GridGrindPlan plan =
        GridGrindPlan.newWorkbook()
            .planId(new PlanId("budget-plan"))
            .saveAs(
                tempDir.resolve("budget.xlsx"), WorkbookPlan.WorkbookPersistence.IfExists.REJECT)
            .journal(ExecutionJournalLevel.VERBOSE)
            .mutate(Targets.sheet(new SheetName("Budget")).ensureExists())
            .mutate(
                Targets.cell(new SheetName("Budget"), new CellAddress("A1"))
                    .set(Values.text("Owner")))
            .inspect(Targets.cell(new SheetName("Budget"), new CellAddress("A1")).read())
            .assertThat(
                Targets.cell(new SheetName("Budget"), new CellAddress("A1"))
                    .valueEquals(ExpectedValues.text("Owner")));
    WorkbookPlan canonical = plan.toPlan();
    assertEquals("budget-plan", canonical.planId().orElseThrow().value());
    assertEquals("mutation-001", canonical.steps().get(0).stepId().value());
    assertEquals("mutation-002", canonical.steps().get(1).stepId().value());
    assertEquals("inspection-001", canonical.steps().get(2).stepId().value());
    assertEquals("assertion-001", canonical.steps().get(3).stepId().value());
    WorkbookPlan reread =
        GridGrindJson.readRequest(
            dev.erst.gridgrind.json.codec.GridGrindJsonOutput.writeRequestBytes(plan.toPlan()));
    assertEquals(canonical, reread);
    assertTrue(
        new String(
                dev.erst.gridgrind.json.codec.GridGrindJsonOutput.writeRequestBytes(plan.toPlan()),
                StandardCharsets.UTF_8)
            .contains("\"planId\":\"budget-plan\""));
    ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
    dev.erst.gridgrind.json.codec.GridGrindJsonOutput.writeRequest(
        outputStream, plan.toPlan(), false);
    assertTrue(
        outputStream.toString(StandardCharsets.UTF_8).contains("\"planId\":\"budget-plan\""));
  }

  @Test
  void preservesImportedCanonicalExecutionFieldsAtTheBoundary() {
    WorkbookPlan canonical =
        new WorkbookPlan(
            dev.erst.gridgrind.core.plan.GridGrindProtocolVersion.current(),
            Optional.of(new PlanId("canonical-plan")),
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.journal(
                new dev.erst.gridgrind.core.plan.ExecutionJournalInput(
                    ExecutionJournalLevel.NORMAL)),
            FormulaEnvironmentInput.empty(),
            List.of());

    WorkbookPlan roundTrip = GridGrindPlan.from(canonical).toPlan();
    assertEquals(canonical, roundTrip);
  }

  @Test
  void importedPlansContinueAutoGeneratedCountersAndRespectNamedSteps() {
    WorkbookPlan imported =
        WorkbookPlan.identified(
            dev.erst.gridgrind.core.plan.GridGrindProtocolVersion.current(),
            new PlanId("imported"),
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    new StepId("mutation-001"),
                    Targets.sheet(new SheetName("Budget")).selector(),
                    new dev.erst.gridgrind.core.plan.WorkbookMutationAction.EnsureSheet()),
                new InspectionStep(
                    new StepId("inspection-001"),
                    Targets.cell(new SheetName("Budget"), new CellAddress("A1")).selector(),
                    new dev.erst.gridgrind.core.plan.SheetIntrospectionQuery.GetCells()),
                new AssertionStep(
                    new StepId("assertion-001"),
                    Targets.cell(new SheetName("Budget"), new CellAddress("A1")).selector(),
                    new dev.erst.gridgrind.core.plan.CellAssertion.CellValue(
                        ExpectedValues.toCellScalarValue(ExpectedValues.blank())))));

    WorkbookPlan plan =
        GridGrindPlan.from(imported)
            .clearPlanId()
            .mutate(Targets.sheet(new SheetName("Budget")).renameTo(new SheetName("Budget 2026")))
            .inspect(Targets.cell(new SheetName("Budget 2026"), new CellAddress("A1")).read())
            .assertThat(
                Targets.cell(new SheetName("Budget 2026"), new CellAddress("A1"))
                    .valueEquals(ExpectedValues.blank()))
            .mutate(
                Targets.sheet(new SheetName("Budget 2026"))
                    .ensureExists()
                    .named(new StepId("keep-sheet")))
            .toPlan();

    assertEquals("mutation-002", plan.steps().get(3).stepId().value());
    assertEquals("inspection-002", plan.steps().get(4).stepId().value());
    assertEquals("assertion-002", plan.steps().get(5).stepId().value());
    assertEquals("keep-sheet", plan.steps().get(6).stepId().value());
  }

  @Test
  void conveniencePersistenceMethodsCoverInMemoryAndOverwriteBranches() {
    GridGrindPlan plan = GridGrindPlan.newWorkbook();
    assertInstanceOf(
        WorkbookPlan.WorkbookPersistence.None.class,
        plan.saveAs(tempDir.resolve("copy.xlsx"), WorkbookPlan.WorkbookPersistence.IfExists.REJECT)
            .inMemoryOnly()
            .toPlan()
            .persistence());
    assertInstanceOf(
        WorkbookPlan.WorkbookPersistence.Overwrite.class,
        GridGrindPlan.open(tempDir.resolve("source.xlsx"))
            .overwrite(OoxmlPersistenceSecurityInput.none())
            .toPlan()
            .persistence());
    assertEquals(
        OoxmlPersistenceSecurityInput.none(),
        ((WorkbookPlan.WorkbookPersistence.SaveAs)
                GridGrindPlan.open(tempDir.resolve("source.xlsx"))
                    .saveAs(
                        tempDir.resolve("copy.xlsx"),
                        WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                        OoxmlPersistenceSecurityInput.none())
                    .toPlan()
                    .persistence())
            .security()
            .orElseThrow());
    assertThrows(
        IllegalStateException.class,
        () ->
            GridGrindPlan.open(tempDir.resolve("source.xlsx"))
                .saveAs(
                    tempDir.resolve("copy.xlsx"),
                    WorkbookPlan.WorkbookPersistence.IfExists.REJECT));
  }

  @Test
  void runsOneTableAwareOfficeWorkflowWithExplicitExecutor() throws Exception {
    Files.writeString(tempDir.resolve("item.txt"), "Hosting");
    Path outputPath = tempDir.resolve("budget.xlsx");

    GridGrindPlan plan =
        GridGrindPlan.newWorkbook()
            .saveAs(outputPath, WorkbookPlan.WorkbookPersistence.IfExists.REPLACE)
            .journal(ExecutionJournalLevel.VERBOSE)
            .mutate(Targets.sheet(new SheetName("Budget")).ensureExists())
            .mutate(
                Targets.range(new SheetName("Budget"), "A1:B3")
                    .setRows(
                        List.of(
                            Values.row(Values.text("Item"), Values.text("Amount")),
                            Values.row(Values.text("Hosting"), Values.number(100.0)),
                            Values.row(Values.text("Travel"), Values.number(50.0)))))
            .mutate(
                Targets.tableOnSheet("BudgetTable", new SheetName("Budget"))
                    .define(
                        Tables.define(
                            "BudgetTable",
                            new SheetName("Budget"),
                            "A1:B3",
                            false,
                            Tables.noStyle())))
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
                    .valueEquals(ExpectedValues.number(125.0)));

    WorkbookResult.Success response =
        assertInstanceOf(
            WorkbookResult.Success.class,
            GridGrindEngine.requestExecutor()
                .execute(
                    plan.toPlan(),
                    new GridGrindRequestInputs(
                        tempDir,
                        tempDir.resolve("temp-root"),
                        new GridGrindExecutionGrant.Bounded(
                            List.of(
                                new GridGrindExecutionGrant.ReadAuthority.File(
                                    tempDir.resolve("item.txt"))),
                            List.of(
                                "ENSURE_SHEET",
                                "SET_RANGE",
                                "SET_TABLE",
                                "SET_CELL",
                                "GET_CELLS",
                                "EXPECT_CELL_VALUE"),
                            new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
                            new GridGrindExecutionGrant.PublicationAuthority.SaveAs(
                                outputPath, WorkbookPlan.WorkbookPersistence.IfExists.REPLACE),
                            List.of(),
                            dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum())),
                    GridGrindProgressSink.NOOP));

    assertTrue(Files.exists(outputPath));
    assertEquals(1, response.assertions().size());
    assertNotNull(response.journal());
    SheetInspectionResult.CellsResult cellsResult =
        assertInstanceOf(
            SheetInspectionResult.CellsResult.class, response.inspections().getFirst());
    dev.erst.gridgrind.core.model.value.CellReport.NumberReport numberReport =
        assertInstanceOf(
            dev.erst.gridgrind.core.model.value.CellReport.NumberReport.class,
            cellsResult.cells().getFirst());
    assertEquals("Budget", cellsResult.sheetName().value());
    assertEquals("B2", numberReport.address().value());
    assertEquals(125.0, numberReport.numberValue().orElseThrow());
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
        import dev.erst.gridgrind.core.result.WorkbookResult;
        import java.nio.file.Path;

        public final class JavaAuthoringWorkflowExampleBridge {
          private JavaAuthoringWorkflowExampleBridge() {}

          public static WorkbookResult run(Path workspace) throws Exception {
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
      WorkbookResult.Success response =
          assertInstanceOf(WorkbookResult.Success.class, runMethod.invoke(null, tempDir));
      assertTrue(Files.exists(tempDir.resolve("budget.xlsx")));
      assertFalse(Files.exists(tempDir.resolve(".gridgrind")));
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
