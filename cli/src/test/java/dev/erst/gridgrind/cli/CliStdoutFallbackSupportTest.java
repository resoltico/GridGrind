package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.cli.discovery.CliFailureReport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused coverage for stdout-fallback helper overloads. */
class CliStdoutFallbackSupportTest extends GridGrindCliTestSupport {
  @Test
  void convenienceOverloadsSerializeTheirDefaultCompactPayloads() throws IOException {
    CliFailureReport failureReport =
        new CliFailureReport(
            GridGrindProtocolVersion.current(),
            2,
            "parse-arguments",
            "parse-arguments",
            GridGrindProblemCode.INVALID_ARGUMENTS,
            "bad flag",
            Optional.empty(),
            Optional.of("--flag"),
            List.of("gridgrind --help"),
            Optional.of("Use one primary command."));
    GridGrindResponse response =
        GridGrindResponses.success(java.util.List.of(), java.util.List.of(), java.util.List.of());
    RequestDoctorReport doctorReport =
        RequestDoctorReport.clean(
            new RequestDoctorReport.Summary(
                "NEW", "NONE", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 0, 0, 0, 0));

    CliStdoutFallbackSupport.StdoutFallback failureFallback =
        CliStdoutFallbackSupport.cliFailureReport("CLI failure report", failureReport);
    CliStdoutFallbackSupport.StdoutFallback responseFallback =
        CliStdoutFallbackSupport.response("response", response);
    CliStdoutFallbackSupport.StdoutFallback doctorFallback =
        CliStdoutFallbackSupport.doctorReport("doctor report", doctorReport);

    assertEquals("CLI failure report", failureFallback.description());
    assertEquals(failureReport, cliFailure(failureFallback.payload()));
    assertEquals(
        GridGrindResponse.Success.class,
        GridGrindJson.readResponse(responseFallback.payload()).getClass());
    assertEquals(doctorReport, GridGrindJson.readRequestDoctorReport(doctorFallback.payload()));
  }
}
