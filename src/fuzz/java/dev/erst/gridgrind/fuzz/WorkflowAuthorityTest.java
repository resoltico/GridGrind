package dev.erst.gridgrind.fuzz;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.core.discovery.RecipeRegistry;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.engine.runtime.DefaultGridGrindRequestExecutor;
import dev.erst.gridgrind.engine.runtime.ExecutionInputBindings;
import dev.erst.gridgrind.fuzz.support.FuzzWorkbookIo;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Prevents workflow fuzzing from succeeding merely by being denied before execution. */
class WorkflowAuthorityTest {
  @Test
  void fixtureAuthorityReachesActualWorkbookExecution(@TempDir Path root) {
    var request = RecipeRegistry.recipeFor("BUDGET").orElseThrow().plan();
    var executor = new DefaultGridGrindRequestExecutor();
    var denied =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            executor.execute(
                request,
                new ExecutionInputBindings(
                    root, root.resolve("denied-tmp"), ExecutionGrant.denied())));
    assertEquals(GridGrindProblemCode.AUTHORITY_DENIED, denied.problem().code());
    assertInstanceOf(
        WorkbookResult.Success.class,
        executor.execute(request, FuzzWorkbookIo.executionBindings(root, request)));
  }
}
