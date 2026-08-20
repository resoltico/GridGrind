package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused coverage for explicit execution-root path and workbook helpers. */
class ExecutionPathCoverageTest {
  @Test
  void typedExecutionPathHelpersResolveAgainstTheProvidedWorkingDirectory() {
    WorkbookPlan request =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("input.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None(),
            dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
            dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
            List.of());
    Path workingDirectory = Path.of("/tmp/gridgrind-explicit-root");

    assertEquals(
        workingDirectory.resolve("input.xlsx").normalize().toString(),
        ExecutionRequestPaths.reqSourcePath(request, workingDirectory));
    assertEquals(
        workingDirectory.resolve("input.xlsx").normalize(),
        ExecutionRequestPaths.normalizePath("input.xlsx", workingDirectory));
  }

  @Test
  void workbookOpenUsesTheProvidedWorkingDirectoryForNewSources() throws IOException {
    Path workingDirectory = Files.createTempDirectory("gridgrind-open-workdir-");
    ExecutionWorkbookSupport workbookSupport =
        ExecutionContextFixtureSupport.workbookSupport(workingDirectory);
    ExecutionInputBindings bindings =
        new ExecutionInputBindings(workingDirectory, workingDirectory.resolve("scratch"));
    try (RequestPathAccess access =
            new RequestPathAccess(workingDirectory, bindings.tempFileFactory());
        var workbook =
            workbookSupport.openWorkbook(
                new WorkbookPlan.WorkbookSource.New(),
                null,
                bindings.withRequestPathAccess(access))) {
      assertNotNull(workbook);
    }
  }

  @Test
  void typedExecutionPathHelpersExposeWorkbookAndPersistenceReferences() {
    WorkbookPlan existingSaveAsRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("input.xlsx"),
            new WorkbookPlan.WorkbookPersistence.SaveAs(
                "out.xlsx",
                WorkbookPlan.WorkbookPersistence.IfExists.REJECT,
                dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput.none()),
            dev.erst.gridgrind.contract.dto.ExecutionPolicyInput.defaults(),
            dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput.empty(),
            List.of());
    WorkbookPlan overwriteRequest =
        WorkbookPlan.standard(
            new WorkbookPlan.WorkbookSource.ExistingFile("input.xlsx"),
            new WorkbookPlan.WorkbookPersistence.Overwrite(
                dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput.none()),
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
            .Overwrite("/tmp/gridgrind/input.xlsx"),
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
  void symlinkEscapingWorkingDirectoryIsRejected() throws IOException {
    Path workDir = Files.createTempDirectory("gridgrind-symlink-test");
    Path outside = Files.createTempDirectory("gridgrind-symlink-outside");
    Path symlink = workDir.resolve("escape");
    Files.createSymbolicLink(symlink, outside);
    try {
      org.junit.jupiter.api.Assertions.assertThrows(
          UnsafePathAccessException.class,
          () -> RequestPathBinding.bindExistingRead("escape/secret.xlsx", workDir));
    } finally {
      Files.delete(symlink);
      Files.delete(outside);
      Files.delete(workDir);
    }
  }

  @Test
  void symlinkWithinWorkingDirectoryIsRejected() throws IOException {
    Path workDir = Files.createTempDirectory("gridgrind-symlink-internal-test");
    Path subDir = Files.createTempDirectory(workDir, "subdir");
    Path symlink = workDir.resolve("internal-link");
    Files.createSymbolicLink(symlink, subDir);
    try {
      org.junit.jupiter.api.Assertions.assertThrows(
          UnsafePathAccessException.class,
          () -> RequestPathBinding.bindWriteTarget("internal-link/file.xlsx", workDir));
    } finally {
      Files.delete(symlink);
      Files.delete(subDir);
      Files.delete(workDir);
    }
  }

  @Test
  void danglingSymlinkInWorkingDirectoryIsRejected() throws IOException {
    Path workDir = Files.createTempDirectory("gridgrind-symlink-dangling-test");
    Path nonExistent = workDir.resolve("gone");
    Path symlink = workDir.resolve("dangler");
    Files.createSymbolicLink(symlink, nonExistent);
    try {
      org.junit.jupiter.api.Assertions.assertThrows(
          UnsafePathAccessException.class,
          () -> RequestPathBinding.bindWriteTarget("dangler/file.xlsx", workDir));
    } finally {
      Files.delete(symlink);
      Files.delete(workDir);
    }
  }

  @Test
  void relativePathTraversalEscapingWorkingDirectoryIsRejected() {
    Path workingDirectory = Path.of("/tmp/gridgrind");

    IllegalArgumentException traversalFailure =
        org.junit.jupiter.api.Assertions.assertThrows(
            IllegalArgumentException.class,
            () -> ExecutionRequestPaths.normalizePath("../../etc/passwd", workingDirectory));
    assertTrue(traversalFailure.getMessage().contains("../../etc/passwd"));

    // Sibling traversal that stays within working directory is allowed
    assertEquals(
        workingDirectory.resolve("subdir/workbook.xlsx").normalize(),
        ExecutionRequestPaths.normalizePath("subdir/workbook.xlsx", workingDirectory));

    RequestPathEscapeException absoluteFailure =
        org.junit.jupiter.api.Assertions.assertThrows(
            RequestPathEscapeException.class,
            () ->
                ExecutionRequestPaths.normalizePath("/tmp/other/workbook.xlsx", workingDirectory));
    assertTrue(absoluteFailure.getMessage().contains("/tmp/other/workbook.xlsx"));

    assertEquals(
        workingDirectory.resolve("subdir/workbook.xlsx").normalize(),
        ExecutionRequestPaths.normalizePath(
            workingDirectory.resolve("subdir/workbook.xlsx").toString(), workingDirectory));
  }

  @Test
  void sourceFilePathTraversalEscapingWorkingDirectoryIsRejected() throws IOException {
    Path workDir = Files.createTempDirectory("gridgrind-source-path-test");
    try (RequestPathAccess access =
        new RequestPathAccess(
            workDir, (prefix, suffix) -> Files.createTempFile(workDir, prefix, suffix))) {
      org.junit.jupiter.api.Assertions.assertThrows(
          RequestPathEscapeException.class,
          () ->
              access.materializeRead("../../etc/passwd", "test input", "gridgrind-test-", ".txt"));
    } finally {
      Files.delete(workDir);
    }
  }

  @Test
  void sourceFilePathWithinWorkingDirectoryIsAllowed() throws IOException {
    Path workDir = Files.createTempDirectory("gridgrind-source-path-allowed-test");
    Path file = Files.createTempFile(workDir, "input", ".txt");
    Files.writeString(file, "inside root");
    try (RequestPathAccess access =
        new RequestPathAccess(
            workDir, (prefix, suffix) -> Files.createTempFile(workDir, prefix, suffix))) {
      assertEquals(
          "inside root",
          Files.readString(
              access.materializeRead(
                  file.getFileName().toString(), "test input", "gridgrind-test-", ".txt")));
    } finally {
      Files.delete(file);
      Files.delete(workDir);
    }
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
        ExecutionContextFixtureSupport.startJournal(request, ExecutionJournalSink.NOOP);
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

    WorkbookResult.Failure failure =
        ExecutionResponseSupport.failureResponse(
            GridGrindProtocolVersion.V2, journal, request, problem, 1, "step-1");

    assertEquals(CalculationReport.notRequested(), failure.calculation());
    assertEquals(problem, failure.problem());
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    ExecutionJournal.Outcome.Failed outcome =
        assertInstanceOf(ExecutionJournal.Outcome.Failed.class, failure.journal().outcome());
    assertEquals(1, outcome.failedStep().orElseThrow().failedStepIndex());
    assertEquals("step-1", outcome.failedStep().orElseThrow().failedStepId());
  }
}
