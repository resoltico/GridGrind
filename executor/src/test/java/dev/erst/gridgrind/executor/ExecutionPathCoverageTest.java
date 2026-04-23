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
}
