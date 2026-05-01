package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.query.InspectionResult;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Shared support logic that keeps the public response DTO file focused on contract shapes. */
final class GridGrindResponseSupport {
  private GridGrindResponseSupport() {}

  static GridGrindResponse.Success success(
      GridGrindProtocolVersion protocolVersion,
      GridGrindResponsePersistence.PersistenceOutcome persistence,
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections) {
    return new GridGrindResponse.Success(
        protocolVersionOrCurrent(protocolVersion),
        syntheticSuccessJournal(),
        CalculationReport.notRequested(),
        Objects.requireNonNull(persistence, "persistence must not be null"),
        copyValues(Objects.requireNonNull(warnings, "warnings must not be null"), "warnings"),
        copyValues(Objects.requireNonNull(assertions, "assertions must not be null"), "assertions"),
        copyValues(
            Objects.requireNonNull(inspections, "inspections must not be null"), "inspections"));
  }

  static GridGrindResponse.Failure failure(
      GridGrindProtocolVersion protocolVersion, GridGrindProblemDetail.Problem problem) {
    Objects.requireNonNull(problem, "problem must not be null");
    return new GridGrindResponse.Failure(
        protocolVersionOrCurrent(protocolVersion),
        syntheticFailureJournal(problem.code()),
        CalculationReport.notRequested(),
        problem);
  }

  static ExecutionJournal syntheticSuccessJournal() {
    return syntheticJournal(ExecutionJournal.Status.SUCCEEDED, Optional.empty());
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
    if (status == ExecutionJournal.Status.FAILED && normalizedFailureCode.isEmpty()) {
      throw new IllegalArgumentException("failureCode must be present when status is FAILED");
    }
    if (status != ExecutionJournal.Status.FAILED && normalizedFailureCode.isPresent()) {
      throw new IllegalArgumentException("failureCode is only permitted when status is FAILED");
    }
    return new ExecutionJournal(
        Optional.empty(),
        ExecutionJournalLevel.NORMAL,
        new ExecutionJournal.SourceSummary(Optional.empty(), Optional.empty()),
        new ExecutionJournal.PersistenceSummary(Optional.empty(), Optional.empty()),
        ExecutionJournal.Phase.notStarted(),
        ExecutionJournal.Phase.notStarted(),
        ExecutionJournal.Phase.notStarted(),
        new ExecutionJournal.Calculation(
            ExecutionJournal.Phase.notStarted(), ExecutionJournal.Phase.notStarted()),
        ExecutionJournal.Phase.notStarted(),
        ExecutionJournal.Phase.notStarted(),
        List.of(),
        List.of(),
        new ExecutionJournal.Outcome(
            status, 0, 0, 0, Optional.empty(), Optional.empty(), normalizedFailureCode),
        List.of());
  }

  static GridGrindProtocolVersion protocolVersionOrCurrent(
      GridGrindProtocolVersion protocolVersion) {
    return protocolVersion == null ? GridGrindProtocolVersion.current() : protocolVersion;
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
    if (causes == null) {
      return List.of();
    }
    List<GridGrindProblemDetail.ProblemCause> copy = new ArrayList<>(causes.size());
    for (GridGrindProblemDetail.ProblemCause cause : causes) {
      copy.add(Objects.requireNonNull(cause, "causes must not contain nulls"));
    }
    return List.copyOf(copy);
  }
}
