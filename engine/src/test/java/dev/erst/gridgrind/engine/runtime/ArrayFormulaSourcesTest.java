package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.mutate;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.mutations;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.request;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.model.formula.ArrayFormulaSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Exercises declared external array-formula sources through publication and readback. */
class ArrayFormulaSourcesTest {
  @TempDir Path directory;

  @Test
  void materializesUtf8FileAndPublishesNormalizedArrayFormula() throws IOException {
    Files.writeString(directory.resolve("formula.txt"), "{=SUM(1,2)}", StandardCharsets.UTF_8);
    WorkbookPlan plan = plan(TextSourceSpec.utf8File("formula.txt"));
    assertFalse(SourceBackedPlanResolver.requiresStandardInput(plan));
    WorkbookResult result =
        new DefaultGridGrindRequestExecutor()
            .execute(plan, ExecutionInputBindingsFixtureSupport.bindings(directory, plan));
    assertInstanceOf(WorkbookResult.Success.class, result, result::toString);
    assertPublishedFormula();
  }

  @Test
  void detectsAndMaterializesStandardInputArrayFormula() throws IOException {
    WorkbookPlan plan = plan(TextSourceSpec.standardInput());
    assertTrue(SourceBackedPlanResolver.requiresStandardInput(plan));
    WorkbookResult result =
        new DefaultGridGrindRequestExecutor()
            .execute(
                plan,
                ExecutionInputBindingsFixtureSupport.bindings(
                    directory, plan, "{=SUM(1,2)}".getBytes(StandardCharsets.UTF_8)));
    assertInstanceOf(WorkbookResult.Success.class, result, result::toString);
    assertPublishedFormula();
  }

  @Test
  void invalidMaterializedArrayFormulaRetainsItsSourceLocationAndDoesNotPublish() {
    WorkbookPlan plan = plan(TextSourceSpec.standardInput());
    WorkbookResult.Failure failure =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            new DefaultGridGrindRequestExecutor()
                .execute(
                    plan,
                    ExecutionInputBindingsFixtureSupport.bindings(
                        directory, plan, "{= }".getBytes(StandardCharsets.UTF_8))));
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    ProblemContext.ResolveInputs context =
        assertInstanceOf(ProblemContext.ResolveInputs.class, failure.problem().context());
    assertEquals(
        "steps[1].action.formula.source.type",
        context.json().orElseThrow().jsonPathValue().orElseThrow());
    assertTrue(
        failure.problem().message().contains("must not be blank after stripping formula syntax"));
    assertFalse(Files.exists(directory.resolve("output.xlsx")));
  }

  private WorkbookPlan plan(TextSourceSpec source) {
    return request(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.SaveAs(
            "output.xlsx", WorkbookPlan.WorkbookPersistence.IfExists.REJECT),
        mutations(
            mutate(
                new SheetSelector.ByName(new SheetName("Budget")),
                new WorkbookMutationAction.EnsureSheet()),
            mutate(
                new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("B1:B2")),
                new CellMutationAction.SetArrayFormula(new ArrayFormulaSpec(source)))));
  }

  private void assertPublishedFormula() throws IOException {
    Path output = directory.resolve("output.xlsx");
    assertTrue(Files.isRegularFile(output));
    try (var workbook = ExecutionContextFixtureSupport.openWorkbook(output)) {
      var cell =
          assertInstanceOf(
              ExcelCellSnapshot.FormulaSnapshot.class,
              workbook.sheet("Budget").cells().snapshotCell("B1"));
      assertEquals("SUM(1,2)", cell.formula());
    }
  }
}
