package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence.Preservation;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookLocation;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Request-scoped mutable facts for one full-XSSF execution; never shared beyond that execution. */
final class FullXssfWorkflowState {
  private final WorkbookWorkflowExecutionContext executionContext;
  private final WorkbookLocation workbookLocation;
  private final List<AssertionResult> hostAssertions = new ArrayList<>();
  private final List<Preservation> hostPreservation = new ArrayList<>();
  private final CollectedAssertionFailures collectedAssertionFailures =
      new CollectedAssertionFailures();
  private final FormulaOriginTracker formulaOrigins = new FormulaOriginTracker();
  private CalculationReport calculation;
  private boolean calculationExecuted;
  private HostAcceptanceSnapshot hostAcceptanceSnapshot = HostAcceptanceSnapshot.empty();

  FullXssfWorkflowState(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExcelWorkbook workbook,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      ExecutionInputBindings bindings) {
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(warnings, "warnings must not be null");
    Objects.requireNonNull(journal, "journal must not be null");
    Objects.requireNonNull(bindings, "bindings must not be null");
    List<AssertionResult> planAssertions = new ArrayList<>();
    List<InspectionResult> inspections = new ArrayList<>();
    executionContext =
        new WorkbookWorkflowExecutionContext(
            protocolVersion,
            request,
            workbook,
            journal,
            warnings,
            planAssertions,
            hostAssertions,
            hostPreservation,
            inspections);
    workbookLocation =
        ExecutionRequestPaths.workbookLocationFor(
            request.source(), request.persistence(), bindings.workingDirectory());
    calculation = CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy());
  }

  WorkbookWorkflowExecutionContext executionContext() {
    return executionContext;
  }

  WorkbookLocation workbookLocation() {
    return workbookLocation;
  }

  List<AssertionResult> hostAssertions() {
    return hostAssertions;
  }

  List<Preservation> hostPreservation() {
    return hostPreservation;
  }

  CollectedAssertionFailures collectedAssertionFailures() {
    return collectedAssertionFailures;
  }

  FormulaOriginTracker formulaOrigins() {
    return formulaOrigins;
  }

  CalculationReport calculation() {
    return calculation;
  }

  void calculation(CalculationReport calculation) {
    this.calculation = Objects.requireNonNull(calculation, "calculation must not be null");
  }

  boolean calculationExecuted() {
    return calculationExecuted;
  }

  void calculationExecuted(boolean calculationExecuted) {
    this.calculationExecuted = calculationExecuted;
  }

  HostAcceptanceSnapshot hostAcceptanceSnapshot() {
    return hostAcceptanceSnapshot;
  }

  void hostAcceptanceSnapshot(HostAcceptanceSnapshot hostAcceptanceSnapshot) {
    this.hostAcceptanceSnapshot =
        Objects.requireNonNull(hostAcceptanceSnapshot, "hostAcceptanceSnapshot must not be null");
  }
}
