package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.query.InspectionResult;
import java.util.List;

/** Public response factories that keep {@link GridGrindResponse} focused on contract shapes. */
public final class GridGrindResponses {
  private GridGrindResponses() {}

  /**
   * Creates a successful response with a synthetic success journal and a not-requested calculation
   * report.
   */
  public static GridGrindResponse.Success success(
      GridGrindProtocolVersion protocolVersion,
      GridGrindResponse.PersistenceOutcome persistence,
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections) {
    return GridGrindResponseSupport.success(
        protocolVersion, persistence, warnings, assertions, inspections);
  }

  /** Creates one successful response using the current protocol version and no persistence. */
  public static GridGrindResponse.Success success(
      List<RequestWarning> warnings,
      List<AssertionResult> assertions,
      List<InspectionResult> inspections) {
    return success(
        GridGrindProtocolVersion.current(),
        new GridGrindResponse.PersistenceOutcome.NotSaved(),
        warnings,
        assertions,
        inspections);
  }

  /**
   * Creates a failure response with a synthetic failure journal and a not-requested calculation
   * report.
   */
  public static GridGrindResponse.Failure failure(
      GridGrindProtocolVersion protocolVersion, GridGrindResponse.Problem problem) {
    return GridGrindResponseSupport.failure(protocolVersion, problem);
  }

  /** Creates one failed response using the current protocol version. */
  public static GridGrindResponse.Failure failure(GridGrindResponse.Problem problem) {
    return failure(GridGrindProtocolVersion.current(), problem);
  }
}
