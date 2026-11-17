package dev.erst.gridgrind.fuzz;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.core.discovery.RecipeRegistry;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.engine.api.GridGrindEngine;
import dev.erst.gridgrind.engine.api.GridGrindRequestInputs;
import dev.erst.gridgrind.fuzz.support.FuzzWorkbookIo;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Prevents workflow fuzzing from succeeding merely by being denied before execution. */
class WorkflowAuthorityTest {
  @Test
  void fixtureAuthorityReachesActualWorkbookExecution(@TempDir Path root) {
    var request = RecipeRegistry.recipeFor("BUDGET").orElseThrow().plan();
    var executor = GridGrindEngine.requestExecutor();
    var denied =
        assertInstanceOf(
            WorkbookResult.Failure.class,
            executor.execute(
                request,
                new GridGrindRequestInputs(
                    root, root.resolve("denied-tmp"), ExecutionGrant.denied())));
    assertEquals(GridGrindProblemCode.AUTHORITY_DENIED, denied.problem().code());
    assertInstanceOf(
        WorkbookResult.Success.class,
        executor.execute(request, FuzzWorkbookIo.executionInputs(root, request)));
  }
}
