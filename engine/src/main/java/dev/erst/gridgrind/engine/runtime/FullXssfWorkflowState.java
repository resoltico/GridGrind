package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.ResolvedPlan;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.RequestWarning;
import dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Preservation;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookLocation;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import java.util.List;
import java.util.Objects;

/** Request-scoped mutable facts for one full-XSSF execution; never shared beyond that execution. */
final class FullXssfWorkflowState {
  private final WorkbookWorkflowExecutionContext executionContext;
  private final ExecutionSession session;
  private final ResolvedPlan plan;
  private final WorkbookLocation workbookLocation;
  private final List<AssertionResult> hostAssertions;
  private final List<Preservation> hostPreservation;
  private final CollectedAssertionFailures collectedAssertionFailures =
      new CollectedAssertionFailures();
  private final FormulaOriginTracker formulaOrigins = new FormulaOriginTracker();
  private CalculationReport calculation;
  private boolean calculationExecuted;
  private HostAcceptanceSnapshot hostAcceptanceSnapshot = HostAcceptanceSnapshot.empty();

  FullXssfWorkflowState(
      ProtocolVersion protocolVersion,
      ResolvedPlan plan,
      XssfWorkbook workbook,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal,
      ExecutionInputBindings bindings) {
    this.session = bindings.session();
    this.plan = Objects.requireNonNull(plan, "plan must not be null");
    WorkbookPlan request = plan.request();
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(warnings, "warnings must not be null");
    Objects.requireNonNull(journal, "journal must not be null");
    Objects.requireNonNull(bindings, "bindings must not be null");
    hostAssertions = bindings.session().hostAssertions;
    hostPreservation = bindings.session().preservation;
    List<AssertionResult> planAssertions = bindings.session().planAssertions;
    List<InspectionResult> inspections = bindings.session().inspections;
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

  ResolvedPlan plan() {
    return plan;
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
    session.calculation = java.util.Optional.of(calculation);
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
