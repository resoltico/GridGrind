package dev.erst.gridgrind.fuzz;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.engine.backend.xssf.SheetNotFoundException;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import dev.erst.gridgrind.fuzz.support.FuzzWorkbookIo;
import dev.erst.gridgrind.fuzz.support.GridGrindFuzzData;
import dev.erst.gridgrind.fuzz.support.OperationSequenceModel;
import dev.erst.gridgrind.fuzz.support.XlsxRoundTripVerifier;
import java.io.IOException;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

/** Retains the known outcomes formerly asserted through the deleted replay-operator DTOs. */
class RoundTripRegressionTest {
  @Test
  void namedRangeShiftRetainsTheMissingSheetBoundary() throws IOException {
    var commands =
        OperationSequenceModel.nextWorkbookCommands(
            GridGrindFuzzData.replay(input("named_range_shift_overwrite_invalid.bin")));
    try (var workbook = XssfWorkbooks.create()) {
      var failure =
          assertThrows(
              SheetNotFoundException.class,
              () -> new WorkbookExecutionEngine().apply(workbook, commands));
      assertEquals("Sheet does not exist: W", failure.getMessage());
    }
  }

  @Test
  void partialCollapsedColumnUngroupRemainsAValidRoundTrip(@TempDir Path root) throws IOException {
    var commands =
        OperationSequenceModel.nextWorkbookCommands(
            GridGrindFuzzData.replay(
                input("partial_collapsed_column_ungroup_roundtrip_success.bin")));
    try (var workbook = XssfWorkbooks.create()) {
      new WorkbookExecutionEngine().apply(workbook, commands);
      Path output = root.resolve("grouping.xlsx");
      FuzzWorkbookIo.saveWorkbook(workbook, output);
      XlsxRoundTripVerifier.requireRoundTripReadable(workbook, output, commands);
    }
  }

  private static byte[] input(String name) throws IOException {
    try (var input =
        RoundTripRegressionTest.class.getResourceAsStream("/seeds/xlsx-roundtrip/" + name)) {
      if (input == null) throw new IOException("Missing retained regression input: " + name);
      return input.readAllBytes();
    }
  }
}
