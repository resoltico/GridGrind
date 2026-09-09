package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.util.List;
import java.util.Objects;

/** Request workflow implementations for workbook, direct-event, and streaming execution modes. */
final class ExecutionWorkflowSupport {
  private final FullXssfWorkflow fullXssfWorkflow;
  private final ExecutionResponseSupport responseSupport;
  private final ExecutionDirectEventReadWorkflow directEventReadWorkflow;
  private final ExecutionStreamingWorkflow streamingWorkflow;

  ExecutionWorkflowSupport(
      ExecutionWorkbookSupport workbookSupport,
      ExecutionCalculationSupport calculationSupport,
      ExecutionStepSupport stepSupport,
      ExecutionResponseSupport responseSupport,
      TempFileFactory tempFileFactory) {
    ExecutionWorkbookSupport requiredWorkbookSupport =
        Objects.requireNonNull(workbookSupport, "workbookSupport must not be null");
    ExecutionCalculationSupport requiredCalculationSupport =
        Objects.requireNonNull(calculationSupport, "calculationSupport must not be null");
    ExecutionStepSupport requiredStepSupport =
        Objects.requireNonNull(stepSupport, "stepSupport must not be null");
    this.responseSupport =
        Objects.requireNonNull(responseSupport, "responseSupport must not be null");
    FullXssfStepExecution fullXssfStepExecution =
        new FullXssfStepExecution(requiredStepSupport, this.responseSupport);
    this.fullXssfWorkflow =
        new FullXssfWorkflow(
            requiredCalculationSupport,
            fullXssfStepExecution,
            new FullXssfCalculationCheckpoint(requiredCalculationSupport, this.responseSupport),
            new FullXssfPersistence(requiredWorkbookSupport, this.responseSupport),
            new HostAcceptanceExecutor(requiredStepSupport),
            this.responseSupport);
    TempFileFactory requiredTempFileFactory =
        Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
    this.directEventReadWorkflow =
        new ExecutionDirectEventReadWorkflow(
            requiredStepSupport, this.responseSupport, requiredTempFileFactory);
    this.streamingWorkflow =
        new ExecutionStreamingWorkflow(
            requiredWorkbookSupport,
            requiredCalculationSupport,
            requiredStepSupport,
            requiredTempFileFactory);
  }

  WorkbookResult executeWorkbookWorkflow(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExcelWorkbook workbook,
      ExecutionModeInput executionMode,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      ExecutionInputBindings bindings) {
    return fullXssfWorkflow.execute(
        protocolVersion, request, workbook, executionMode, warnings, journal, bindings);
  }

  WorkbookResult executeDirectEventReadWorkflow(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      ExecutionInputBindings bindings) {
    return directEventReadWorkflow.execute(protocolVersion, request, warnings, journal, bindings);
  }

  WorkbookResult executeStreamingWorkflow(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionModeInput executionMode,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      ExecutionInputBindings bindings) {
    return streamingWorkflow.execute(
        protocolVersion, request, executionMode, warnings, journal, bindings);
  }
}
