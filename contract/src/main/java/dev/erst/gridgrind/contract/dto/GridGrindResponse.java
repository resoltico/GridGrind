package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail.Problem;
import dev.erst.gridgrind.contract.dto.GridGrindResponsePersistence.PersistenceOutcome;
import dev.erst.gridgrind.contract.query.InspectionResult;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * Structured protocol response for both successful workbook workflows and deterministic failures.
 */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "status")
@JsonSubTypes({
  @JsonSubTypes.Type(value = GridGrindResponse.Success.class, name = "SUCCESS"),
  @JsonSubTypes.Type(value = GridGrindResponse.Failure.class, name = "ERROR")
})
public sealed interface GridGrindResponse {
  /** Protocol version negotiated for this response. */
  GridGrindProtocolVersion protocolVersion();

  /** Structured execution journal captured for this run, even when it failed. */
  ExecutionJournal journal();

  /**
   * Structured calculation report captured for this run, even when calculation was not requested.
   */
  CalculationReport calculation();

  /** Successful workbook execution result. */
  record Success(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournal journal,
      CalculationReport calculation,
      PersistenceOutcome persistence,
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections)
      implements GridGrindResponse {
    public Success {
      Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
      Objects.requireNonNull(journal, "journal must not be null");
      Objects.requireNonNull(calculation, "calculation must not be null");
      Objects.requireNonNull(persistence, "persistence must not be null");
      warnings =
          GridGrindResponseSupport.copyValues(
              Objects.requireNonNull(warnings, "warnings must not be null"), "warnings");
      assertions =
          GridGrindResponseSupport.copyValues(
              Objects.requireNonNull(assertions, "assertions must not be null"), "assertions");
      inspections = GridGrindResponseSupport.copyValues(inspections, "inspections");
    }
  }

  /** Failed workbook execution with a structured problem. */
  record Failure(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournal journal,
      CalculationReport calculation,
      Problem problem)
      implements GridGrindResponse {
    public Failure {
      Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
      Objects.requireNonNull(journal, "journal must not be null");
      Objects.requireNonNull(calculation, "calculation must not be null");
      Objects.requireNonNull(problem, "problem must not be null");
    }
  }

  /** Creates a synthetic success journal for non-step-oriented responses. */
  static ExecutionJournal syntheticSuccessJournal() {
    return GridGrindResponseSupport.syntheticSuccessJournal();
  }

  /** Creates a synthetic failed journal for non-step-oriented responses. */
  static ExecutionJournal syntheticFailureJournal(GridGrindProblemCode failureCode) {
    return GridGrindResponseSupport.syntheticFailureJournal(failureCode);
  }

  /** Creates a synthetic journal for non-step-oriented responses with explicit failure state. */
  static ExecutionJournal syntheticJournal(
      ExecutionJournal.Status status, Optional<GridGrindProblemCode> failureCode) {
    return GridGrindResponseSupport.syntheticJournal(status, failureCode);
  }
}
