package dev.erst.gridgrind.fuzz.engine;

import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import dev.erst.gridgrind.fuzz.support.GridGrindFuzzData;
import dev.erst.gridgrind.fuzz.support.OperationSequenceModel;
import dev.erst.gridgrind.fuzz.support.WorkbookInvariantChecks;
import java.io.IOException;
import java.util.List;

/** Fuzzes workbook-core command sequences directly against the engine entrypoint. */
public final class CommandSequenceHarness {
  private CommandSequenceHarness() {}

  public static void check(GridGrindFuzzData fuzzData) throws IOException {
    List<ResolvedStep> commands;
    try {
      commands = OperationSequenceModel.nextWorkbookCommands(fuzzData);
    } catch (IllegalArgumentException expected) {
      return;
    }
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      WorkbookExecutionEngine executor = new WorkbookExecutionEngine();

      try {
        executor.apply(workbook, commands);
        WorkbookInvariantChecks.requireWorkbookShape(workbook);
      } catch (IllegalArgumentException expected) {
        // Invalid ordered sequences are expected to fail within documented validation families.
      }
    }
  }
}
