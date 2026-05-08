package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;

import java.io.IOException;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct coverage for workbook-execution facade overloads. */
class WorkbookExecutionEngineCoverageTest {
  @Test
  void readsFromIterableCommandsUsingUnsavedWorkbookContext() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookExecutionEngine engine = new WorkbookExecutionEngine();

      assertEquals(List.of(), engine.read(workbook, List.of()));
    }
  }
}
