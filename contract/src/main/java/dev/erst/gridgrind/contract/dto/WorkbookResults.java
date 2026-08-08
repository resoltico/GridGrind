package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.query.InspectionResult;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Public result factories that keep {@link WorkbookResult} focused on contract shapes. */
public final class WorkbookResults {
  private WorkbookResults() {}

  /**
   * Creates a successful response with a synthetic success journal and a not-requested calculation
   * report.
   */
  public static WorkbookResult.Success success(
      GridGrindProtocolVersion protocolVersion,
      WorkbookResultPersistence.PersistenceOutcome persistence,
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections) {
    return WorkbookResultSupport.success(
        protocolVersion, persistence, warnings, assertions, inspections);
  }

  /** Creates one successful response using the current protocol version and no persistence. */
  public static WorkbookResult.Success success(
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections) {
    return success(
        GridGrindProtocolVersion.current(),
        new WorkbookResultPersistence.PersistenceOutcome.NotSaved(),
        warnings,
        assertions,
        inspections);
  }

  /**
   * Creates a failure response with a synthetic failure journal and a not-requested calculation
   * report.
   */
  public static WorkbookResult.Failure failure(
      GridGrindProtocolVersion protocolVersion,
      WorkbookResultPersistence.PersistenceOutcome persistence,
      GridGrindProblemDetail.Problem problem) {
    return failure(protocolVersion, Optional.empty(), persistence, problem);
  }

  /** Creates a failure result with the request's optional stable plan identifier. */
  public static WorkbookResult.Failure failure(
      GridGrindProtocolVersion protocolVersion,
      Optional<String> planId,
      WorkbookResultPersistence.PersistenceOutcome persistence,
      GridGrindProblemDetail.Problem problem) {
    return WorkbookResultSupport.failure(protocolVersion, planId, persistence, problem);
  }

  /**
   * Creates a failure response with a synthetic failure journal and a not-requested calculation
   * report.
   */
  public static WorkbookResult.Failure failure(
      GridGrindProtocolVersion protocolVersion, GridGrindProblemDetail.Problem problem) {
    return failure(
        protocolVersion, new WorkbookResultPersistence.PersistenceOutcome.NotSaved(), problem);
  }

  /** Creates one failed response using the current protocol version. */
  public static WorkbookResult.Failure failure(GridGrindProblemDetail.Problem problem) {
    return failure(GridGrindProtocolVersion.current(), problem);
  }

  /**
   * Describes persistence that was requested but did not write an artifact.
   *
   * <p>This keeps every failure response truthful about the caller's save intent without inventing
   * an execution path for an artifact that was never written.
   */
  public static WorkbookResultPersistence.PersistenceOutcome unwrittenPersistenceOutcome(
      WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.None _ ->
          new WorkbookResultPersistence.PersistenceOutcome.NotSaved();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          new WorkbookResultPersistence.PersistenceOutcome.SavedAs(
              saveAs.path(), new WorkbookResultPersistence.WriteResult.NotWritten());
      case WorkbookPlan.WorkbookPersistence.Overwrite _ -> {
        Optional<String> sourcePath =
            switch (request.source()) {
              case WorkbookPlan.WorkbookSource.New _ -> Optional.empty();
              case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
                  Optional.of(existingFile.path());
            };
        yield new WorkbookResultPersistence.PersistenceOutcome.Overwritten(
            sourcePath, new WorkbookResultPersistence.WriteResult.NotWritten());
      }
    };
  }

  /**
   * Reclassifies a completed execution as failed when a later execution-channel operation fails.
   *
   * <p>The journal remains the truthful workbook-execution record: writing the CLI response is a
   * transport operation outside the workbook journal. The result status and singular problem report
   * that post-execution failure without discarding completed warnings, assertions, or inspections.
   */
  public static WorkbookResult.Failure afterExecutionFailure(
      WorkbookResult response, GridGrindProblemDetail.Problem problem) {
    Objects.requireNonNull(response, "response must not be null");
    Objects.requireNonNull(problem, "problem must not be null");
    return switch (response) {
      case WorkbookResult.Success success ->
          new WorkbookResult.Failure(
              success.protocolVersion(),
              success.planId(),
              success.journal(),
              success.calculation(),
              success.persistence(),
              success.warnings(),
              success.assertions(),
              success.inspections(),
              problem);
      case WorkbookResult.Failure failure ->
          new WorkbookResult.Failure(
              failure.protocolVersion(),
              failure.planId(),
              failure.journal(),
              failure.calculation(),
              failure.persistence(),
              failure.warnings(),
              failure.assertions(),
              failure.inspections(),
              problem);
    };
  }
}
