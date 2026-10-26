package dev.erst.gridgrind.core.result;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.GridGrindProtocolVersion;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.problem.DiagnosticOrder;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail.Problem;
import dev.erst.gridgrind.core.problem.RequestWarning;
import dev.erst.gridgrind.core.result.WorkbookResultPersistence.PersistenceOutcome;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Structured protocol result emitted only after workbook execution begins. */
@Union(discriminator = "status")
public sealed interface WorkbookResult {
  /** Protocol version negotiated for this response. */
  GridGrindProtocolVersion protocolVersion();

  /**
   * Optional caller-supplied plan identifier, owned by the result rather than execution telemetry.
   */
  Optional<PlanId> planId();

  /** Structured execution journal captured for this run, even when it failed. */
  ExecutionJournal journal();

  /**
   * Structured calculation report captured for this run, even when calculation was not requested.
   */
  CalculationReport calculation();

  /** Structured workbook-persistence outcome returned for every response. */
  PersistenceOutcome persistence();

  /** Successful workbook execution result. */
  @Variant(id = "SUCCEEDED", summary = "Successful workbook execution result.")
  record Success(
      GridGrindProtocolVersion protocolVersion,
      Optional<PlanId> planId,
      ExecutionJournal journal,
      CalculationReport calculation,
      PersistenceOutcome persistence,
      WorkbookExecutionEvidence evidence,
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections)
      implements WorkbookResult {
    public Success {
      Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
      Objects.requireNonNull(planId, "planId must not be null");
      Objects.requireNonNull(journal, "journal must not be null");
      Objects.requireNonNull(calculation, "calculation must not be null");
      Objects.requireNonNull(persistence, "persistence must not be null");
      Objects.requireNonNull(evidence, "evidence must not be null");
      warnings =
          DiagnosticOrder.warnings(
              WorkbookResultSupport.copyValues(
                  Objects.requireNonNull(warnings, "warnings must not be null"), "warnings"));
      assertions =
          WorkbookResultSupport.copyValues(
              Objects.requireNonNull(assertions, "assertions must not be null"), "assertions");
      inspections = WorkbookResultSupport.copyValues(inspections, "inspections");
    }
  }

  /** Failed workbook execution with a structured problem. */
  @Variant(id = "FAILED", summary = "Failed workbook execution with a structured problem.")
  record Failure(
      GridGrindProtocolVersion protocolVersion,
      Optional<PlanId> planId,
      ExecutionJournal journal,
      CalculationReport calculation,
      PersistenceOutcome persistence,
      WorkbookExecutionEvidence evidence,
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections,
      Problem problem)
      implements WorkbookResult {
    public Failure {
      Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
      Objects.requireNonNull(planId, "planId must not be null");
      Objects.requireNonNull(journal, "journal must not be null");
      Objects.requireNonNull(calculation, "calculation must not be null");
      Objects.requireNonNull(persistence, "persistence must not be null");
      Objects.requireNonNull(evidence, "evidence must not be null");
      warnings =
          DiagnosticOrder.warnings(
              WorkbookResultSupport.copyValues(
                  Objects.requireNonNull(warnings, "warnings must not be null"), "warnings"));
      assertions =
          WorkbookResultSupport.copyValues(
              Objects.requireNonNull(assertions, "assertions must not be null"), "assertions");
      inspections =
          WorkbookResultSupport.copyValues(
              Objects.requireNonNull(inspections, "inspections must not be null"), "inspections");
      Objects.requireNonNull(problem, "problem must not be null");
    }
  }

  /** Creates a synthetic success journal for non-step-oriented responses. */
  static ExecutionJournal syntheticSuccessJournal() {
    return WorkbookResultSupport.syntheticSuccessJournal();
  }

  /** Creates a synthetic failed journal for non-step-oriented responses. */
  static ExecutionJournal syntheticFailureJournal(GridGrindProblemCode failureCode) {
    return WorkbookResultSupport.syntheticFailureJournal(failureCode);
  }

  /** Creates a synthetic journal for non-step-oriented responses with explicit failure state. */
  static ExecutionJournal syntheticJournal(
      ExecutionJournal.Status status, Optional<GridGrindProblemCode> failureCode) {
    return WorkbookResultSupport.syntheticJournal(status, failureCode);
  }
}
