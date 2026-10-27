package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.IOException;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct coverage for workbook-execution facade overloads. */
class WorkbookExecutionEngineCoverageTest {
  @Test
  void readsFromIterableCommandsUsingUnsavedWorkbookContext() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      WorkbookExecutionEngine engine = new WorkbookExecutionEngine();

      assertEquals(List.of(), engine.read(workbook, List.of()));
    }
  }
}
