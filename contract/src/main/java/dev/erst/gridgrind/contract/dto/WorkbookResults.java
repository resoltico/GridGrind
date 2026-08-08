package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.query.InspectionResult;
import java.util.List;
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
}
