package dev.erst.gridgrind.jazzer.engine;

import com.code_intelligence.jazzer.api.FuzzedDataProvider;
import com.code_intelligence.jazzer.junit.FuzzTest;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.jazzer.support.HarnessTelemetry;
import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import dev.erst.gridgrind.jazzer.support.OperationSequenceModel;
import dev.erst.gridgrind.jazzer.support.SequenceIntrospection;
import dev.erst.gridgrind.jazzer.support.WorkbookInvariantChecks;
import dev.erst.gridgrind.jazzer.support.XlsxRoundTripVerifier;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/** Fuzzes `.xlsx` save and reopen invariants after bounded command sequences. */
class XlsxRoundTripFuzzTest {
  private static final HarnessTelemetry TELEMETRY =
      HarnessTelemetry.forHarness(JazzerHarness.XLSX_ROUND_TRIP);

  @FuzzTest
  void roundTrip(FuzzedDataProvider data) throws IOException {
    TELEMETRY.beginIteration(data.remainingBytes());
    List<WorkbookCommand> commands;
    try {
      commands = OperationSequenceModel.nextWorkbookCommands(data);
    } catch (IllegalArgumentException expected) {
      TELEMETRY.recordGeneratedInvalid();
      return;
    }
    TELEMETRY.recordSequenceKinds(SequenceIntrospection.commandKinds(commands));
    TELEMETRY.recordStyleKinds(SequenceIntrospection.styleKindsFromCommands(commands));
    Path directory = Files.createTempDirectory("gridgrind-jazzer-roundtrip-");
    Path workbookPath = directory.resolve("workbook.xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();
      try {
        executor.apply(workbook, commands);
        WorkbookInvariantChecks.requireWorkbookShape(workbook);
        workbook.save(workbookPath);
        XlsxRoundTripVerifier.requireRoundTripReadable(workbookPath, commands);
        TELEMETRY.recordSuccess();
      } catch (IllegalArgumentException expected) {
        // Invalid command sequences are expected to fail before or during persistence.
        TELEMETRY.recordExpectedInvalid(expected);
      } catch (IOException unexpected) {
        TELEMETRY.recordUnexpectedFailure(unexpected);
        throw unexpected;
      } catch (RuntimeException unexpected) {
        TELEMETRY.recordUnexpectedFailure(unexpected);
        throw unexpected;
      }
    } catch (IOException unexpected) {
      TELEMETRY.recordUnexpectedFailure(unexpected);
      throw unexpected;
    }
  }
}
