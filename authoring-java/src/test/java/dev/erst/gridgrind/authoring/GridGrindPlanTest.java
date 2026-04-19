package dev.erst.gridgrind.authoring;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.executor.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.executor.ExecutionInputBindings;
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

/** End-to-end tests for the Java-first authoring API introduced in Phase 8. */
class GridGrindPlanTest {
  @TempDir Path tempDir;

  @Test
  void serializesOneFluentPlanBackToTheCanonicalWorkbookPlan() throws Exception {
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
  }

  @Test
  void runsOneTableAwareOfficeWorkflowIncludingSourceBackedRowSelection() throws Exception {
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
                        new TableInput(
                            "BudgetTable", "Budget", "A1:B3", false, new TableStyleInput.None())))
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
            plan.run(
                new DefaultGridGrindRequestExecutor(),
                new ExecutionInputBindings(tempDir, (byte[]) null)));

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
    Path classesDirectory = Files.createDirectories(tempDir.resolve("example-classes"));
    DiagnosticCollector<JavaFileObject> diagnostics = new DiagnosticCollector<>();

    try (StandardJavaFileManager fileManager =
        compiler.getStandardFileManager(null, Locale.ROOT, StandardCharsets.UTF_8)) {
      Iterable<? extends JavaFileObject> compilationUnits =
          fileManager.getJavaFileObjects(examplePath.toFile());
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
