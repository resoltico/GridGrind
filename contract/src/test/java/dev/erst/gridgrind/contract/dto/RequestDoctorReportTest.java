package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for machine-readable request doctor reports. */
class RequestDoctorReportTest {
  @Test
  void reportFactoriesNormalizeSeverityAndCopyWarnings() {
    RequestDoctorReport.Summary summary =
        new RequestDoctorReport.Summary(
            "NEW", "NONE", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 0, 0, 0, 0);
    RequestWarning warning =
        new RequestWarning(
            dev.erst.gridgrind.contract.dto.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
            0,
            "step-1",
            "SET_CELL",
            "warning");
    List<RequestWarning> mutableWarnings = new java.util.ArrayList<>(List.of(warning));
    GridGrindProblemDetail.Problem problem =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST,
            "bad request",
            new ProblemContext.ValidateRequest(
                ProblemContextRequestSurfaces.RequestShape.known("NEW", "NONE")));

    RequestDoctorReport clean = RequestDoctorReport.clean(summary);
    RequestDoctorReport warnings = RequestDoctorReport.warnings(summary, mutableWarnings);
    RequestDoctorReport invalid = RequestDoctorReport.invalid(summary, mutableWarnings, problem);
    RequestDoctorReport invalidBatch =
        RequestDoctorReport.invalid(summary, mutableWarnings, List.of(problem));
    mutableWarnings.add(
        new RequestWarning(
            dev.erst.gridgrind.contract.dto.GridGrindWarningCode.UNQUOTED_SHEET_NAME_IN_FORMULA,
            1,
            "step-2",
            "SET_RANGE",
            "ignored"));

    assertEquals(AnalysisSeverity.INFO, clean.severity());
    assertEquals(AnalysisSeverity.WARNING, warnings.severity());
    assertEquals(AnalysisSeverity.ERROR, invalid.severity());
    assertEquals(List.of(problem), invalid.problems());
    assertEquals(List.of(problem), invalidBatch.problems());
    assertEquals(Optional.empty(), clean.primaryProblem());
    assertEquals(Optional.of(problem), invalid.primaryProblem());
    assertEquals(List.of(warning), warnings.warnings());
    assertEquals(GridGrindProtocolVersion.current(), clean.protocolVersion());
    assertThrows(UnsupportedOperationException.class, () -> warnings.warnings().add(warning));
  }

  @Test
  void reportAndSummaryValidationRejectInconsistentShapes() {
    RequestDoctorReport.Summary summary =
        new RequestDoctorReport.Summary(
            "NEW", "NONE", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 0, 0, 0, 0);
    GridGrindProblemDetail.Problem problem =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST,
            "bad request",
            new ProblemContext.ValidateRequest(
                ProblemContextRequestSurfaces.RequestShape.known("NEW", "NONE")));

    assertEquals(
        "valid doctor report requires a summary",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport(
                        GridGrindProtocolVersion.current(),
                        AnalysisSeverity.INFO,
                        true,
                        Optional.empty(),
                        List.of(),
                        List.of()))
            .getMessage());
    assertEquals(
        "valid doctor report cannot contain problems",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport(
                        GridGrindProtocolVersion.current(),
                        AnalysisSeverity.INFO,
                        true,
                        Optional.of(summary),
                        List.of(),
                        List.of(problem)))
            .getMessage());
    assertEquals(
        "invalid doctor report requires at least one problem",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport(
                        GridGrindProtocolVersion.current(),
                        AnalysisSeverity.ERROR,
                        false,
                        Optional.of(summary),
                        List.of(),
                        List.of()))
            .getMessage());
    assertEquals(
        "clean valid doctor report requires INFO severity",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport(
                        GridGrindProtocolVersion.current(),
                        AnalysisSeverity.WARNING,
                        true,
                        Optional.of(summary),
                        List.of(),
                        List.of()))
            .getMessage());
    assertEquals(
        "invalid doctor report requires ERROR severity",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport(
                        GridGrindProtocolVersion.current(),
                        AnalysisSeverity.WARNING,
                        false,
                        Optional.of(summary),
                        List.of(),
                        List.of(problem)))
            .getMessage());
    assertEquals(
        "valid doctor report warnings require WARNING severity",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport(
                        GridGrindProtocolVersion.current(),
                        AnalysisSeverity.INFO,
                        true,
                        Optional.of(summary),
                        List.of(
                            new RequestWarning(
                                dev.erst.gridgrind.contract.dto.GridGrindWarningCode
                                    .UNQUOTED_SHEET_NAME_IN_FORMULA,
                                0,
                                "step-1",
                                "SET_CELL",
                                "warning")),
                        List.of()))
            .getMessage());
    assertEquals(
        "mutationStepCount + assertionStepCount + inspectionStepCount must equal stepCount",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport.Summary(
                        "NEW", "NONE", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 1, 0, 0, 0))
            .getMessage());
    assertEquals(
        "sourceType must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new RequestDoctorReport.Summary(
                        null, "NONE", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 0, 0, 0, 0))
            .getMessage());
    assertEquals(
        "stepCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport.Summary(
                        "NEW", "NONE", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, -1, 0, 0, 0))
            .getMessage());
    assertEquals(
        "sourceType must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport.Summary(
                        " ", "NONE", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 0, 0, 0, 0))
            .getMessage());
  }

  @Test
  void supportsExplicitProtocolVersionAndRejectsNullWarnings() {
    GridGrindProblemDetail.Problem problem =
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST,
            "bad request",
            new ProblemContext.ValidateRequest(
                ProblemContextRequestSurfaces.RequestShape.known("NEW", "NONE")));
    List<RequestWarning> warningsWithNull = new java.util.ArrayList<>();
    warningsWithNull.add(null);

    assertEquals(
        "warnings must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new RequestDoctorReport(
                        GridGrindProtocolVersion.V2,
                        AnalysisSeverity.ERROR,
                        false,
                        Optional.empty(),
                        null,
                        List.of(problem)))
            .getMessage());
    assertEquals(
        "warnings must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    new RequestDoctorReport(
                        GridGrindProtocolVersion.current(),
                        AnalysisSeverity.WARNING,
                        true,
                        Optional.of(
                            new RequestDoctorReport.Summary(
                                "NEW",
                                "NONE",
                                "FULL_XSSF",
                                "DO_NOT_CALCULATE",
                                false,
                                false,
                                0,
                                0,
                                0,
                                0)),
                        warningsWithNull,
                        List.of()))
            .getMessage());
  }
}
