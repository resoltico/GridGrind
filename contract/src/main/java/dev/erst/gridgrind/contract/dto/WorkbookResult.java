package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail.Problem;
import dev.erst.gridgrind.contract.dto.WorkbookResultPersistence.PersistenceOutcome;
import dev.erst.gridgrind.contract.query.InspectionResult;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Structured protocol result emitted only after workbook execution begins. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "status")
@JsonSubTypes({
  @JsonSubTypes.Type(value = WorkbookResult.Success.class, name = "SUCCEEDED"),
  @JsonSubTypes.Type(value = WorkbookResult.Failure.class, name = "FAILED")
})
public sealed interface WorkbookResult {
  /** Protocol version negotiated for this response. */
  GridGrindProtocolVersion protocolVersion();

  /** Optional caller-supplied plan identifier, owned by the result rather than execution telemetry. */
  Optional<String> planId();

  /** Structured execution journal captured for this run, even when it failed. */
  ExecutionJournal journal();

  /**
   * Structured calculation report captured for this run, even when calculation was not requested.
   */
  CalculationReport calculation();

  /** Structured workbook-persistence outcome returned for every response. */
  PersistenceOutcome persistence();

  /** Successful workbook execution result. */
  record Success(
      GridGrindProtocolVersion protocolVersion,
      @com.fasterxml.jackson.annotation.JsonInclude(com.fasterxml.jackson.annotation.JsonInclude.Include.NON_ABSENT)
          Optional<String> planId,
      ExecutionJournal journal,
      CalculationReport calculation,
      PersistenceOutcome persistence,
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections)
      implements WorkbookResult {
    public Success {
      Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
      planId = WorkbookResultSupport.optionalPlanId(planId);
      Objects.requireNonNull(journal, "journal must not be null");
      Objects.requireNonNull(calculation, "calculation must not be null");
      Objects.requireNonNull(persistence, "persistence must not be null");
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
  record Failure(
      GridGrindProtocolVersion protocolVersion,
      @com.fasterxml.jackson.annotation.JsonInclude(com.fasterxml.jackson.annotation.JsonInclude.Include.NON_ABSENT)
          Optional<String> planId,
      ExecutionJournal journal,
      CalculationReport calculation,
      PersistenceOutcome persistence,
      List<AssertionResult> assertions,
      Problem problem)
      implements WorkbookResult {
    public Failure {
      Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
      planId = WorkbookResultSupport.optionalPlanId(planId);
      Objects.requireNonNull(journal, "journal must not be null");
      Objects.requireNonNull(calculation, "calculation must not be null");
      Objects.requireNonNull(persistence, "persistence must not be null");
      assertions =
          WorkbookResultSupport.copyValues(
              Objects.requireNonNull(assertions, "assertions must not be null"), "assertions");
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
