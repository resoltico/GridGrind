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
            "NEW", "NONE", "FULL_XSSF", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 0, 0, 0, 0);
    RequestWarning warning = new RequestWarning(0, "step-1", "SET_CELL", "warning");
    List<RequestWarning> mutableWarnings = new java.util.ArrayList<>(List.of(warning));
    GridGrindResponse.Problem problem =
        GridGrindResponse.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST,
            "bad request",
            new ProblemContext.ValidateRequest(ProblemContext.RequestShape.known("NEW", "NONE")));

    RequestDoctorReport clean = RequestDoctorReport.clean(summary);
    RequestDoctorReport warnings = RequestDoctorReport.warnings(summary, mutableWarnings);
    RequestDoctorReport invalid = RequestDoctorReport.invalid(summary, mutableWarnings, problem);
    mutableWarnings.add(new RequestWarning(1, "step-2", "SET_RANGE", "ignored"));

    assertEquals(AnalysisSeverity.INFO, clean.severity());
    assertEquals(AnalysisSeverity.WARNING, warnings.severity());
    assertEquals(AnalysisSeverity.ERROR, invalid.severity());
    assertEquals(List.of(warning), warnings.warnings());
    assertEquals(GridGrindProtocolVersion.current(), clean.protocolVersion());
    assertThrows(UnsupportedOperationException.class, () -> warnings.warnings().add(warning));
  }

  @Test
  void reportAndSummaryValidationRejectInconsistentShapes() {
    RequestDoctorReport.Summary summary =
        new RequestDoctorReport.Summary(
            "NEW", "NONE", "FULL_XSSF", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 0, 0, 0, 0);
    GridGrindResponse.Problem problem =
        GridGrindResponse.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST,
            "bad request",
            new ProblemContext.ValidateRequest(ProblemContext.RequestShape.known("NEW", "NONE")));

    assertEquals(
        "summary must not be null for a valid doctor report",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport(
                        GridGrindProtocolVersion.current(),
                        AnalysisSeverity.INFO,
                        true,
                        Optional.empty(),
                        List.of(),
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "problem must be null for a valid doctor report",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport(
                        GridGrindProtocolVersion.current(),
                        AnalysisSeverity.INFO,
                        true,
                        Optional.of(summary),
                        List.of(),
                        Optional.of(problem)))
            .getMessage());
    assertEquals(
        "problem must not be null for an invalid doctor report",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport(
                        GridGrindProtocolVersion.current(),
                        AnalysisSeverity.ERROR,
                        false,
                        Optional.of(summary),
                        List.of(),
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "severity must be INFO when a valid doctor report contains no warnings",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport(
                        GridGrindProtocolVersion.current(),
                        AnalysisSeverity.WARNING,
                        true,
                        Optional.of(summary),
                        List.of(),
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "severity must be ERROR when a doctor report is invalid",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport(
                        GridGrindProtocolVersion.current(),
                        AnalysisSeverity.WARNING,
                        false,
                        Optional.of(summary),
                        List.of(),
                        Optional.of(problem)))
            .getMessage());
    assertEquals(
        "severity must be WARNING when a valid doctor report contains warnings",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport(
                        GridGrindProtocolVersion.current(),
                        AnalysisSeverity.INFO,
                        true,
                        Optional.of(summary),
                        List.of(new RequestWarning(0, "step-1", "SET_CELL", "warning")),
                        Optional.empty()))
            .getMessage());
    assertEquals(
        "mutationStepCount + assertionStepCount + inspectionStepCount must equal stepCount",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport.Summary(
                        "NEW",
                        "NONE",
                        "FULL_XSSF",
                        "FULL_XSSF",
                        "DO_NOT_CALCULATE",
                        false,
                        false,
                        1,
                        0,
                        0,
                        0))
            .getMessage());
    assertEquals(
        "sourceType must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new RequestDoctorReport.Summary(
                        null,
                        "NONE",
                        "FULL_XSSF",
                        "FULL_XSSF",
                        "DO_NOT_CALCULATE",
                        false,
                        false,
                        0,
                        0,
                        0,
                        0))
            .getMessage());
    assertEquals(
        "stepCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport.Summary(
                        "NEW",
                        "NONE",
                        "FULL_XSSF",
                        "FULL_XSSF",
                        "DO_NOT_CALCULATE",
                        false,
                        false,
                        -1,
                        0,
                        0,
                        0))
            .getMessage());
    assertEquals(
        "sourceType must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RequestDoctorReport.Summary(
                        " ",
                        "NONE",
                        "FULL_XSSF",
                        "FULL_XSSF",
                        "DO_NOT_CALCULATE",
                        false,
                        false,
                        0,
                        0,
                        0,
                        0))
            .getMessage());
  }

  @Test
  void supportsExplicitProtocolVersionAndRejectsNullWarnings() {
    GridGrindResponse.Problem problem =
        GridGrindResponse.Problem.of(
            GridGrindProblemCode.INVALID_REQUEST,
            "bad request",
            new ProblemContext.ValidateRequest(ProblemContext.RequestShape.known("NEW", "NONE")));
    List<RequestWarning> warningsWithNull = new java.util.ArrayList<>();
    warningsWithNull.add(null);

    RequestDoctorReport invalid =
        new RequestDoctorReport(
            GridGrindProtocolVersion.V1,
            AnalysisSeverity.ERROR,
            false,
            Optional.empty(),
            null,
            Optional.of(problem));

    assertEquals(GridGrindProtocolVersion.V1, invalid.protocolVersion());
    assertEquals(List.of(), invalid.warnings());
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
                                "FULL_XSSF",
                                "DO_NOT_CALCULATE",
                                false,
                                false,
                                0,
                                0,
                                0,
                                0)),
                        warningsWithNull,
                        Optional.empty()))
            .getMessage());
  }
}
