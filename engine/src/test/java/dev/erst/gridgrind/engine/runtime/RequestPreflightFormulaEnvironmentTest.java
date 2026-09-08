package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.FormulaExternalWorkbookInput;
import dev.erst.gridgrind.contract.dto.FormulaMissingWorkbookPolicy;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Verifies preflight admission of external-workbook formula environments. */
class RequestPreflightFormulaEnvironmentTest {
  @TempDir Path temporaryDirectory;

  @Test
  void preservesFormulaWorkbookPathEscapeDiagnostics() {
    WorkbookPlan request = plan("../outside.xlsx");

    RequestPreflight.Result result = RequestPreflight.verify(request, bindingsFor(request));

    assertEquals(
        List.of(GridGrindProblemCode.PATH_ESCAPES_ROOT),
        result.problems().stream().map(GridGrindProblemDetail.Problem::code).toList());
  }

  @Test
  void reportsMissingFormulaEnvironmentWorkbooksThroughTheResolvedInputBoundary() {
    WorkbookPlan request = plan("missing-external.xlsx");

    RequestPreflight.Result result = RequestPreflight.verify(request, bindingsFor(request));

    assertFalse(result.preparedRequest().isPresent());
    assertEquals(
        List.of(GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND),
        result.problems().stream().map(GridGrindProblemDetail.Problem::code).toList());
    assertEquals(
        "missing-external.xlsx",
        ((ProblemContext.ResolveInputs) result.problems().getFirst().context())
            .input()
            .inputPathValue()
            .orElseThrow());
  }

  @Test
  void reportsFormulaEnvironmentAuthorityFailuresWithoutLeakingThePrivateBindingPath()
      throws Exception {
    Path externalWorkbook =
        Files.write(temporaryDirectory.resolve("external.xlsx"), new byte[] {1});
    WorkbookPlan request = plan(externalWorkbook.getFileName().toString());
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(
            temporaryDirectory,
            temporaryDirectory.resolve("scratch"),
            ExecutionGrantTestSupport.noPublication());
    List<GridGrindProblemDetail.Problem> problems = new ArrayList<>();

    try (RequestPathAccess access =
        new RequestPathAccess(
            bindings.workingDirectory(), bindings.tempFileFactory(), bindings.executionGrant())) {
      RequestPreflightPaths.verify(request, bindings.withRequestPathAccess(access), problems);
    }

    assertEquals(
        List.of(GridGrindProblemCode.AUTHORITY_DENIED),
        problems.stream().map(GridGrindProblemDetail.Problem::code).toList());
    assertTrue(
        problems.stream()
            .map(GridGrindProblemDetail.Problem::message)
            .noneMatch(message -> message.contains("gridgrind-formula-workbook-")));
  }

  private WorkbookPlan plan(String externalWorkbookPath) {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExecutionPolicyInput.defaults(),
        new FormulaEnvironmentInput(
            List.of(new FormulaExternalWorkbookInput("External.xlsx", externalWorkbookPath)),
            FormulaMissingWorkbookPolicy.ERROR,
            List.of()),
        List.of());
  }

  private ExecutionInputBindings bindingsFor(WorkbookPlan request) {
    return new ExecutionInputBindings(
        temporaryDirectory,
        temporaryDirectory.resolve("scratch"),
        ExecutionGrantTestSupport.forPlan(request, temporaryDirectory));
  }
}
