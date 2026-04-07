package dev.erst.gridgrind.jazzer.engine;

import com.code_intelligence.jazzer.api.FuzzedDataProvider;
import com.code_intelligence.jazzer.junit.FuzzTest;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.jazzer.support.GridGrindFuzzData;
import dev.erst.gridgrind.jazzer.support.HarnessTelemetry;
import dev.erst.gridgrind.jazzer.support.JazzerHarness;
import dev.erst.gridgrind.jazzer.support.OperationSequenceModel;
import dev.erst.gridgrind.jazzer.support.SequenceIntrospection;
import dev.erst.gridgrind.jazzer.support.WorkbookInvariantChecks;
import java.io.IOException;
import java.util.List;

/** Fuzzes workbook-core command sequences directly against the engine entrypoint. */
class WorkbookCommandSequenceFuzzTest {
  private static final HarnessTelemetry TELEMETRY =
      HarnessTelemetry.forHarness(JazzerHarness.ENGINE_COMMAND_SEQUENCE);

  @FuzzTest
  void applyCommands(FuzzedDataProvider data) throws IOException {
    GridGrindFuzzData fuzzData = GridGrindFuzzData.wrap(data);
    TELEMETRY.beginIteration(fuzzData.remainingBytes());
    List<WorkbookCommand> commands;
    try {
      commands = OperationSequenceModel.nextWorkbookCommands(fuzzData);
    } catch (IllegalArgumentException expected) {
      TELEMETRY.recordGeneratedInvalid();
      return;
    }
    TELEMETRY.recordSequenceKinds(SequenceIntrospection.commandKinds(commands));
    TELEMETRY.recordStyleKinds(SequenceIntrospection.styleKindsFromCommands(commands));
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

      try {
        executor.apply(workbook, commands);
        WorkbookInvariantChecks.requireWorkbookShape(workbook);
        TELEMETRY.recordSuccess();
      } catch (IllegalArgumentException expected) {
        // Invalid ordered sequences are expected to fail within documented validation families.
        TELEMETRY.recordExpectedInvalid(expected);
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
