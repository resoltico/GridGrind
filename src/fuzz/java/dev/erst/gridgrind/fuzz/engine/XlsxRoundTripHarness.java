package dev.erst.gridgrind.fuzz.engine;

import dev.erst.gridgrind.core.plan.ResolvedStep;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import dev.erst.gridgrind.fuzz.support.FuzzFilesystem;
import dev.erst.gridgrind.fuzz.support.FuzzWorkbookIo;
import dev.erst.gridgrind.fuzz.support.GridGrindFuzzData;
import dev.erst.gridgrind.fuzz.support.OperationSequenceModel;
import dev.erst.gridgrind.fuzz.support.WorkbookInvariantChecks;
import dev.erst.gridgrind.fuzz.support.XlsxRoundTripVerifier;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/** Fuzzes `.xlsx` save and reopen invariants after bounded command sequences. */
public final class XlsxRoundTripHarness {
  private XlsxRoundTripHarness() {}

  public static void check(GridGrindFuzzData fuzzData) throws IOException {
    List<ResolvedStep> commands;
    try {
      commands = OperationSequenceModel.nextWorkbookCommands(fuzzData);
    } catch (IllegalArgumentException expected) {
      return;
    }
    Path directory = Files.createTempDirectory("gridgrind-jazzer-roundtrip-");
    Path workbookPath = directory.resolve("workbook.xlsx");

    try {
      try (XssfWorkbook workbook = XssfWorkbooks.create()) {
        WorkbookExecutionEngine executor = new WorkbookExecutionEngine();
        try {
          executor.apply(workbook, commands);
          WorkbookInvariantChecks.requireWorkbookShape(workbook);
          FuzzWorkbookIo.saveWorkbook(workbook, workbookPath);
          XlsxRoundTripVerifier.requireRoundTripReadable(workbook, workbookPath, commands);
        } catch (IllegalArgumentException expected) {
          // Invalid command sequences are expected to fail before or during persistence.
        }
      }
    } catch (IOException | RuntimeException | Error failure) {
      FuzzFilesystem.cleanupAfterFailure(directory, failure);
      throw failure;
    }
    FuzzFilesystem.deleteRecursively(directory);
  }
}
