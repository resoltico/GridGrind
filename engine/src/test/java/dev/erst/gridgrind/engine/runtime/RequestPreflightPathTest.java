package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;

import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.FormulaExternalWorkbookInput;
import dev.erst.gridgrind.contract.dto.FormulaMissingWorkbookPolicy;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/**
 * Verifies formula-input and source-path preflight findings independent of general request linting.
 */
class RequestPreflightPathTest {
  @TempDir Path temporaryDirectory;

  @Test
  void batchesFormulaInputsAndPreparesOverwriteTargetsFromExistingSources() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("missing-source.xlsx"),
            new WorkbookPlan.WorkbookPersistence.Overwrite(OoxmlPersistenceSecurityInput.none()),
            ExecutionPolicyInput.defaults(),
            new FormulaEnvironmentInput(
                List.of(
                    new FormulaExternalWorkbookInput(
                        "ExternalOne.xlsx", "missing-external-one.xlsx"),
                    new FormulaExternalWorkbookInput(
                        "ExternalTwo.xlsx", "missing-external-two.xlsx")),
                FormulaMissingWorkbookPolicy.ERROR,
                List.of()),
            List.of());

    RequestPreflight.Result result = verify(request);

    assertFalse(result.preparedRequest().isPresent());
    assertEquals(
        List.of(
            GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND,
            GridGrindProblemCode.INPUT_SOURCE_NOT_FOUND,
            GridGrindProblemCode.WORKBOOK_NOT_FOUND),
        result.problems().stream().map(GridGrindProblemDetail.Problem::code).toList());
  }

  @Test
  void reportsFormulaExternalWorkbookPathsThatEscapeTheExecutionRoot() {
    WorkbookPlan request = escapedFormulaWorkbookPlan();

    RequestPreflight.Result result = verify(request);

    assertEquals(
        List.of(GridGrindProblemCode.PATH_ESCAPES_ROOT),
        result.problems().stream().map(GridGrindProblemDetail.Problem::code).toList());
  }

  @Test
  void recordsRuntimeFormulaExternalPathFailuresDuringPathPreflight() throws Exception {
    WorkbookPlan request = escapedFormulaWorkbookPlan();
    ExecutionInputBindings bindings = bindingsFor(request);
    List<GridGrindProblemDetail.Problem> problems = new java.util.ArrayList<>();

    try (RequestPathAccess access =
        new RequestPathAccess(
            temporaryDirectory, bindings.tempFileFactory(), bindings.executionGrant())) {
      RequestPreflightPaths.verify(request, bindings.withRequestPathAccess(access), problems);
    }

    assertEquals(
        List.of(GridGrindProblemCode.PATH_ESCAPES_ROOT),
        problems.stream().map(GridGrindProblemDetail.Problem::code).toList());
  }

  @Test
  void rejectsColumnEditsAgainstExistingSourcesThatAlreadyContainFormulas() throws Exception {
    Path sourcePath = temporaryDirectory.resolve("existing-formulas.xlsx");
    try (XSSFWorkbook source = new XSSFWorkbook()) {
      source.createSheet("Ops").createRow(0).createCell(0).setCellFormula("1+1");
      try (var output = Files.newOutputStream(sourcePath)) {
        source.write(output);
      }
    }
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile(sourcePath.getFileName().toString()),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.defaults(),
            FormulaEnvironmentInput.empty(),
            List.of(
                new MutationStep(
                    "insert-column",
                    new ColumnBandSelector.Insertion("Ops", 1, 1),
                    new WorkbookMutationAction.InsertColumns())));

    RequestPreflight.Result result = verify(request);

    assertFalse(result.preparedRequest().isPresent());
    GridGrindProblemDetail.Problem problem = result.problems().getFirst();
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, problem.code());
    assertEquals(
        "Column insert, delete, and shift operations must appear before formula authoring because"
            + " formula-bearing workbook column edits are unsupported.",
        problem.message());
    assertInstanceOf(ProblemContext.OpenWorkbook.class, problem.context());
    assertFalse(problem.message().contains("gridgrind-source-workbook-"));
  }

  private RequestPreflight.Result verify(WorkbookPlan request) {
    return RequestPreflight.verify(request, bindingsFor(request));
  }

  private ExecutionInputBindings bindingsFor(WorkbookPlan request) {
    return new ExecutionInputBindings(
        temporaryDirectory,
        temporaryDirectory.resolve("scratch"),
        ExecutionGrantTestSupport.forPlan(request, temporaryDirectory));
  }

  private static WorkbookPlan escapedFormulaWorkbookPlan() {
    return WorkbookPlan.standard(
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExecutionPolicyInput.defaults(),
        new FormulaEnvironmentInput(
            List.of(new FormulaExternalWorkbookInput("Escaped.xlsx", "../escaped.xlsx")),
            FormulaMissingWorkbookPolicy.ERROR,
            List.of()),
        List.of());
  }
}
