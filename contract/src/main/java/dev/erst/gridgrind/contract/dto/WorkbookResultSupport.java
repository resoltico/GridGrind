package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.query.InspectionResult;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Shared support logic that keeps the public result DTO file focused on contract shapes. */
final class WorkbookResultSupport {
  private WorkbookResultSupport() {}

  static WorkbookResult.Success success(
      GridGrindProtocolVersion protocolVersion,
      WorkbookResultPersistence.PersistenceOutcome persistence,
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections) {
    return new WorkbookResult.Success(
        Objects.requireNonNull(protocolVersion, "protocolVersion must not be null"),
        Optional.empty(),
        syntheticSuccessJournal(),
        CalculationReport.notRequested(),
        Objects.requireNonNull(persistence, "persistence must not be null"),
        copyValues(Objects.requireNonNull(warnings, "warnings must not be null"), "warnings"),
        copyValues(Objects.requireNonNull(assertions, "assertions must not be null"), "assertions"),
        copyValues(
            Objects.requireNonNull(inspections, "inspections must not be null"), "inspections"));
  }

  static WorkbookResult.Failure failure(
      GridGrindProtocolVersion protocolVersion,
      Optional<String> planId,
      WorkbookResultPersistence.PersistenceOutcome persistence,
      GridGrindProblemDetail.Problem problem) {
    Objects.requireNonNull(problem, "problem must not be null");
    return new WorkbookResult.Failure(
        Objects.requireNonNull(protocolVersion, "protocolVersion must not be null"),
        optionalPlanId(planId),
        syntheticFailureJournal(problem.code()),
        CalculationReport.notRequested(),
        Objects.requireNonNull(persistence, "persistence must not be null"),
        List.of(),
        problem);
  }

  static ExecutionJournal syntheticSuccessJournal() {
    return syntheticJournal(ExecutionJournal.Status.SUCCEEDED, Optional.empty());
  }

  static Optional<String> optionalPlanId(Optional<String> planId) {
    Optional<String> normalized = Objects.requireNonNullElseGet(planId, Optional::empty);
    return normalized.map(value -> WorkbookPlan.requireNonBlank(value, "planId"));
  }

  static ExecutionJournal syntheticFailureJournal(GridGrindProblemCode failureCode) {
    return syntheticJournal(
        ExecutionJournal.Status.FAILED,
        Optional.of(Objects.requireNonNull(failureCode, "failureCode must not be null")));
  }

  static ExecutionJournal syntheticJournal(
      ExecutionJournal.Status status, Optional<GridGrindProblemCode> failureCode) {
    Objects.requireNonNull(status, "status must not be null");
    Optional<GridGrindProblemCode> normalizedFailureCode =
        Objects.requireNonNullElseGet(failureCode, Optional::empty);
    validateSyntheticOutcomeRequest(status, normalizedFailureCode);
    ExecutionJournal.Outcome outcome = syntheticOutcome(status, normalizedFailureCode);
    return new ExecutionJournal(
        ExecutionJournalLevel.SUMMARY,
        new ExecutionJournal.SourceSummary(Optional.empty(), Optional.empty()),
        ExecutionJournal.Phase.notStarted(),
        ExecutionJournal.Phase.notStarted(),
        ExecutionJournal.Phase.notStarted(),
        new ExecutionJournal.Calculation(
            ExecutionJournal.Phase.notStarted(), ExecutionJournal.Phase.notStarted()),
        ExecutionJournal.Phase.notStarted(),
        ExecutionJournal.Phase.notStarted(),
        List.of(),
        outcome,
        List.of());
  }

  private static void validateSyntheticOutcomeRequest(
      ExecutionJournal.Status status, Optional<GridGrindProblemCode> failureCode) {
    if (status == ExecutionJournal.Status.FAILED && failureCode.isEmpty()) {
      throw new IllegalArgumentException("FAILED outcomes must include failureCode");
    }
    if (status != ExecutionJournal.Status.FAILED && failureCode.isPresent()) {
      throw new IllegalArgumentException("failureCode is only permitted when status is FAILED");
    }
  }

  private static ExecutionJournal.Outcome syntheticOutcome(
      ExecutionJournal.Status status, Optional<GridGrindProblemCode> failureCode) {
    return switch (status) {
      case SUCCEEDED -> ExecutionJournal.Outcome.succeeded(0, 0, 0);
      case FAILED ->
          ExecutionJournal.Outcome.failed(0, 0, 0, failureCode.orElseThrow(), Optional.empty());
      case NOT_STARTED, NOT_REQUESTED ->
          throw new IllegalArgumentException(
              "synthetic journal outcome does not support " + status);
    };
  }

  static List<String> copyDistinctStrings(List<String> values, String fieldName) {
    List<String> copy = copyStrings(values, fieldName);
    if (copy.size() != new java.util.LinkedHashSet<>(copy).size()) {
      throw new IllegalArgumentException(fieldName + " must not contain duplicates");
    }
    return copy;
  }

  static List<String> validateCommonWorkbookSummaryFields(
      int sheetCount, List<String> sheetNames, int namedRangeCount) {
    if (sheetCount < 0) {
      throw new IllegalArgumentException("sheetCount must not be negative");
    }
    if (namedRangeCount < 0) {
      throw new IllegalArgumentException("namedRangeCount must not be negative");
    }
    List<String> copy = copyDistinctStrings(sheetNames, "sheetNames");
    if (sheetCount != copy.size()) {
      throw new IllegalArgumentException("sheetCount must match sheetNames size");
    }
    for (String sheetName : copy) {
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetNames must not contain blank values");
      }
    }
    return copy;
  }

  static List<String> copyStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<String> copy = new ArrayList<>(values.size());
    for (String value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }

  static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }

  static <T> Optional<List<T>> copyOptionalValues(Optional<List<T>> values, String fieldName) {
    Optional<List<T>> normalized = Objects.requireNonNullElseGet(values, Optional::empty);
    if (normalized.isEmpty()) {
      return Optional.empty();
    }
    List<T> copy = copyValues(normalized.orElseThrow(), fieldName);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    return Optional.of(copy);
  }

  static List<GridGrindProblemDetail.ProblemCause> copyProblemCauses(
      List<GridGrindProblemDetail.ProblemCause> causes) {
    Objects.requireNonNull(causes, "causes must not be null");
    List<GridGrindProblemDetail.ProblemCause> copy = new ArrayList<>(causes.size());
    for (GridGrindProblemDetail.ProblemCause cause : causes) {
      copy.add(Objects.requireNonNull(cause, "causes must not contain nulls"));
    }
    return List.copyOf(copy);
  }
}
