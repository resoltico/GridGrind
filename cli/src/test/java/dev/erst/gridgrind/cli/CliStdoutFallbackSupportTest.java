package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.cli.discovery.CliDiagnostic;
import dev.erst.gridgrind.cli.discovery.CliTransport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused coverage for stdout-fallback helper overloads. */
class CliStdoutFallbackSupportTest extends GridGrindCliTestSupport {
  @Test
  void convenienceOverloadsSerializeTheirDefaultCompactPayloads() throws IOException {
    CliDiagnostic failureReport =
        new CliDiagnostic(
            GridGrindProtocolVersion.current(),
            2,
            "cli",
            List.of("gridgrind --help"),
            GridGrindProblemDetail.Problem.of(
                GridGrindProblemCode.INVALID_ARGUMENTS,
                "bad flag",
                new ProblemContext.ParseArguments(CliArgument.named("--flag"))),
            Optional.of(CliTransport.standardOutput()));
    GridGrindResponse response =
        GridGrindResponses.success(java.util.List.of(), java.util.List.of(), java.util.List.of());
    RequestDoctorReport doctorReport =
        RequestDoctorReport.clean(
            new RequestDoctorReport.Summary(
                "NEW", "NONE", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 0, 0, 0, 0));

    CliStdoutFallbackSupport.StdoutFallback failureFallback =
        CliStdoutFallbackSupport.cliDiagnostic(failureReport);
    CliStdoutFallbackSupport.StdoutFallback responseFallback =
        CliStdoutFallbackSupport.response(response);
    CliStdoutFallbackSupport.StdoutFallback doctorFallback =
        CliStdoutFallbackSupport.doctorReport(doctorReport);

    assertEquals(failureReport, cliDiagnostic(failureFallback.payload()));
    assertEquals(
        GridGrindResponse.Success.class,
        GridGrindJson.readResponse(responseFallback.payload()).getClass());
    assertEquals(doctorReport, GridGrindJson.readRequestDoctorReport(doctorFallback.payload()));
  }

  @Test
  void writeKeepsTheRecoveredStdoutPayloadWhenStderrMirroringFailsAfterwards() throws IOException {
    CliDiagnostic diagnostic =
        new CliDiagnostic(
            GridGrindProtocolVersion.current(),
            1,
            "execute",
            List.of(),
            GridGrindProblemDetail.Problem.of(
                GridGrindProblemCode.IO_ERROR,
                "write failed",
                new ProblemContext.ParseArguments(CliArgument.named("--response"))),
            Optional.of(CliTransport.standardOutput()));
    CliStdoutFallbackSupport.StdoutFallback fallback =
        CliStdoutFallbackSupport.cliDiagnostic(diagnostic);
    ByteArrayOutputStream stdout = new ByteArrayOutputStream();
    try (OutputStream failingStderr =
        new OutputStream() {
          @Override
          public void write(int ignored) throws IOException {
            throw new IOException("stderr unavailable");
          }
        }) {
      assertDoesNotThrow(
          () -> CliStdoutFallbackSupport.write(failingStderr, stdout, diagnostic, fallback, false));
    }
    assertEquals(diagnostic, cliDiagnostic(stdout.toByteArray()));
  }
}
