package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Focused coverage for convenience overloads in execution path and workbook helpers. */
class ExecutionPathCoverageTest {
  @Test
  void noArgExecutionPathHelpersResolveAgainstTheProcessWorkingDirectory() {
    WorkbookPlan request =
        new WorkbookPlan(
            new WorkbookPlan.WorkbookSource.ExistingFile("input.xlsx"),
            new WorkbookPlan.WorkbookPersistence.None(),
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
        new WorkbookPlan(
            new WorkbookPlan.WorkbookSource.ExistingFile("input.xlsx"),
            new WorkbookPlan.WorkbookPersistence.SaveAs("out.xlsx"),
            List.of());
    WorkbookPlan overwriteRequest =
        new WorkbookPlan(
            new WorkbookPlan.WorkbookSource.ExistingFile("input.xlsx"),
            new WorkbookPlan.WorkbookPersistence.OverwriteSource(),
            List.of());
    Path workingDirectory = Path.of("/tmp/gridgrind");

    assertEquals(
        new dev.erst.gridgrind.contract.dto.ProblemContext.WorkbookReference.NewWorkbook(),
        ExecutionRequestPaths.workbookReference(
            new WorkbookPlan(
                new WorkbookPlan.WorkbookSource.New(),
                new WorkbookPlan.WorkbookPersistence.None(),
                List.of()),
            workingDirectory));
    assertEquals(
        new dev.erst.gridgrind.contract.dto.ProblemContext.WorkbookReference.ExistingFile(
            "/tmp/gridgrind/input.xlsx"),
        ExecutionRequestPaths.workbookReference(existingSaveAsRequest, workingDirectory));
    assertEquals(
        new dev.erst.gridgrind.contract.dto.ProblemContext.PersistenceReference.SaveAs(
            "/tmp/gridgrind/out.xlsx"),
        ExecutionRequestPaths.persistenceReference(existingSaveAsRequest, workingDirectory));
    assertEquals(
        new dev.erst.gridgrind.contract.dto.ProblemContext.PersistenceReference.OverwriteSource(
            "/tmp/gridgrind/input.xlsx"),
        ExecutionRequestPaths.persistenceReference(overwriteRequest, workingDirectory));
    assertEquals(
        "persistence reference requires a saving policy",
        org.junit.jupiter.api.Assertions.assertThrows(
                IllegalArgumentException.class,
                () ->
                    ExecutionRequestPaths.persistenceReference(
                        new WorkbookPlan(
                            new WorkbookPlan.WorkbookSource.New(),
                            new WorkbookPlan.WorkbookPersistence.None(),
                            List.of()),
                        workingDirectory))
            .getMessage());
  }
}
