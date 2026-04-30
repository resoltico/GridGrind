package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused coverage for convenience overloads in execution path and workbook helpers. */
class ExecutionPathCoverageTest {
  @Test
  void noArgExecutionPathHelpersResolveAgainstTheProcessWorkingDirectory() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("input.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
            dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
            List.of());

    assertEquals(
        Path.of("").toAbsolutePath().normalize().resolve("input.xlsx").normalize().toString(),
        ExecutionRequestPaths.reqSourcePath(request));
    assertEquals(
        Path.of("").toAbsolutePath().normalize().resolve("input.xlsx").normalize(),
        ExecutionRequestPaths.normalizePath("input.xlsx"));
  }

  @Test
  void workbookOpenOverloadUsesTheDefaultWorkingDirectoryForNewSources() throws IOException {
    ExecutionWorkbookSupport workbookSupport = new ExecutionWorkbookSupport(Files::createTempFile);

    try (var workbook = workbookSupport.openWorkbook(new WorkbookPlan.WorkbookSource.New(), null)) {
      assertNotNull(workbook);
    }
  }

  @Test
  void typedExecutionPathHelpersExposeWorkbookAndPersistenceReferences() {
    WorkbookPlan existingSaveAsRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("input.xlsx"),
            new WorkbookPlan.WorkbookPersistence.SaveAs("out.xlsx"),
            dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
            dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
            List.of());
    WorkbookPlan overwriteRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("input.xlsx"),
            new WorkbookPlan.WorkbookPersistence.OverwriteSource(),
            dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
            dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
            List.of());
    Path workingDirectory = Path.of("/tmp/gridgrind");

    assertEquals(
        new dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.WorkbookReference
            .NewWorkbook(),
        ExecutionRequestPaths.workbookReference(
            WorkbookPlan.standard(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
                dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
                List.of()),
            workingDirectory));
    assertEquals(
        new dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.WorkbookReference
            .ExistingFile("/tmp/gridgrind/input.xlsx"),
        ExecutionRequestPaths.workbookReference(existingSaveAsRequest, workingDirectory));
    assertEquals(
        new dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.PersistenceReference
            .SaveAs("/tmp/gridgrind/out.xlsx"),
        ExecutionRequestPaths.persistenceReference(existingSaveAsRequest, workingDirectory));
    assertEquals(
        new dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.PersistenceReference
            .OverwriteSource("/tmp/gridgrind/input.xlsx"),
        ExecutionRequestPaths.persistenceReference(overwriteRequest, workingDirectory));
    assertEquals(
        "persistence reference requires a saving policy",
        org.junit.jupiter.api.Assertions.assertThrows(
                IllegalArgumentException.class,
                () ->
                    ExecutionRequestPaths.persistenceReference(
                        WorkbookPlan.standard(
                            new WorkbookPlan.WorkbookSource.New(),
                            new WorkbookPlan.WorkbookPersistence.None(),
                            dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
                            dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
                            List.of()),
                        workingDirectory))
            .getMessage());
  }

  @Test
  void failureResponseConvenienceOverloadDefaultsCalculationToNotRequested() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
            dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
            List.of());
    ExecutionJournalRecorder journal =
        ExecutionJournalRecorder.start(request, ExecutionJournalSink.NOOP);
    GridGrindProblemDetail.Problem problem =
        new GridGrindProblemDetail.Problem(
            GridGrindProblemCode.INVALID_REQUEST,
            GridGrindProblemCode.INVALID_REQUEST.category(),
            GridGrindProblemCode.INVALID_REQUEST.recovery(),
            "Invalid request",
            "bad request",
            GridGrindProblemCode.INVALID_REQUEST.resolution(),
            new ProblemContext.ExecuteRequest(ProblemContextRequestSurfaces.RequestShape.unknown()),
            Optional.empty(),
            List.of());

    GridGrindResponse.Failure failure =
        ExecutionResponseSupport.failureResponse(
            GridGrindProtocolVersion.V1, journal, 3, problem, 1, "step-1");

    assertEquals(CalculationReport.notRequested(), failure.calculation());
    assertEquals(problem, failure.problem());
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals(1, failure.journal().outcome().failedStepIndex().orElseThrow());
    assertEquals("step-1", failure.journal().outcome().failedStepId().orElseThrow());
  }
}
