package dev.erst.gridgrind.jazzer.support;

import static dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.*;

import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.jazzer.support.OperationSequenceValueFactory.WorkflowStorage;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Builds bounded protocol requests and workbook command sequences for Jazzer harnesses. */
public final class OperationSequenceModel {
  private OperationSequenceModel() {}

  /** Returns a bounded protocol workflow plus any owned local scratch paths it created. */
  public static GeneratedProtocolWorkflow nextProtocolWorkflow(GridGrindFuzzData data)
      throws IOException {
    Objects.requireNonNull(data, "data must not be null");

    String primarySheet = FuzzDataDecoders.nextSheetName(data, true);
    String secondarySheet = FuzzDataDecoders.nextSheetName(data, true);
    List<ProtocolStepSupport.PendingMutation> mutations = new ArrayList<>();
    List<ProtocolStepSupport.PendingAssertion> assertions = new ArrayList<>();
    List<InspectionStep> inspections = new ArrayList<>();
    String workbookNamedRange = nextNamedRangeName(data, true);
    String sheetNamedRange = nextNamedRangeName(data, true);
    String pivotTableName = nextPivotTableName(data, true);

    int operationCount = data.consumeInt(0, 8);
    for (int index = 0; index < operationCount; index++) {
      mutations.add(
          OperationSequenceMutationFactory.nextMutation(
              data,
              primarySheet,
              secondarySheet,
              workbookNamedRange,
              sheetNamedRange,
              pivotTableName));
    }
    int readCount = data.consumeInt(0, 6);
    for (int index = 0; index < readCount; index++) {
      inspections.add(
          OperationSequenceInspectionFactory.nextInspection(
              data,
              index,
              primarySheet,
              secondarySheet,
              workbookNamedRange,
              sheetNamedRange,
              pivotTableName));
    }
    int assertionCount = data.consumeInt(0, 4);
    for (int index = 0; index < assertionCount; index++) {
      assertions.add(
          OperationSequenceObservationSupport.nextAssertion(
              data,
              index,
              primarySheet,
              secondarySheet,
              workbookNamedRange,
              sheetNamedRange,
              pivotTableName));
    }
    WorkflowStorage workflowStorage = nextWorkflowStorage(primarySheet, secondarySheet, data);
    return new GeneratedProtocolWorkflow(
        ProtocolStepSupport.request(
            workflowStorage.source(),
            workflowStorage.persistence(),
            OperationSequenceSelectorSupport.nextExecutionPolicy(
                data,
                primarySheet,
                secondarySheet,
                OperationSequenceSelectorSupport.validFormulaAddress(data)),
            FormulaEnvironmentInput.empty(),
            mutations,
            assertions,
            inspections),
        List.of(workflowStorage.cleanupRoot()));
  }

  /** Returns a bounded sequence of workbook-core commands. */
  public static List<WorkbookCommand> nextWorkbookCommands(GridGrindFuzzData data) {
    Objects.requireNonNull(data, "data must not be null");

    String primarySheet = FuzzDataDecoders.nextSheetName(data, true);
    String secondarySheet = FuzzDataDecoders.nextSheetName(data, true);
    List<WorkbookCommand> commands = new ArrayList<>();
    String workbookNamedRange = nextNamedRangeName(data, true);
    String sheetNamedRange = nextNamedRangeName(data, true);
    String pivotTableName = nextPivotTableName(data, true);
    int commandCount = data.consumeInt(1, 10);
    for (int index = 0; index < commandCount; index++) {
      commands.add(
          OperationSequenceCommandFactory.nextCommand(
              data,
              primarySheet,
              secondarySheet,
              workbookNamedRange,
              sheetNamedRange,
              pivotTableName));
    }
    return List.copyOf(commands);
  }
}
