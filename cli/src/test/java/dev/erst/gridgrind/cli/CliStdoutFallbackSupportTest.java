package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.cli.discovery.CommandError;
import dev.erst.gridgrind.cli.discovery.CliTransport;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.dto.WorkbookResults;
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
    CommandError failureReport =
        new CommandError(
            GridGrindProtocolVersion.current(),
            2,
            "cli",
            List.of("gridgrind --help"),
            List.of(
                GridGrindProblemDetail.Problem.of(
                    GridGrindProblemCode.INVALID_ARGUMENTS,
                    "bad flag",
                    new ProblemContext.ParseArguments(CliArgument.named("--flag")))),
            Optional.of(CliTransport.standardOutput()));
    WorkbookResult response =
        WorkbookResults.success(java.util.List.of(), java.util.List.of(), java.util.List.of());
    RequestDoctorReport doctorReport =
        RequestDoctorReport.clean(
            new RequestDoctorReport.Summary(
                "NEW", "NONE", "FULL_XSSF", "DO_NOT_CALCULATE", false, false, 0, 0, 0, 0));

    CliStdoutFallbackSupport.StdoutFallback failureFallback =
        CliStdoutFallbackSupport.commandError(failureReport);
    CliStdoutFallbackSupport.StdoutFallback responseFallback =
        CliStdoutFallbackSupport.response(response);
    CliStdoutFallbackSupport.StdoutFallback doctorFallback =
        CliStdoutFallbackSupport.doctorReport(doctorReport);

    assertEquals(failureReport, commandError(failureFallback.payload()));
    assertEquals(
        WorkbookResult.Success.class,
        GridGrindJson.readWorkbookResult(responseFallback.payload()).getClass());
    assertEquals(doctorReport, GridGrindJson.readRequestDoctorReport(doctorFallback.payload()));
  }

  @Test
  void writeKeepsTheRecoveredStdoutPayloadWhenStderrMirroringFailsAfterwards() throws IOException {
    CommandError diagnostic =
        new CommandError(
            GridGrindProtocolVersion.current(),
            1,
            "execute",
            List.of(),
            List.of(
                GridGrindProblemDetail.Problem.of(
                    GridGrindProblemCode.IO_ERROR,
                    "write failed",
                    new ProblemContext.ParseArguments(CliArgument.named("--response")))),
            Optional.of(CliTransport.standardOutput()));
    CliStdoutFallbackSupport.StdoutFallback fallback =
        CliStdoutFallbackSupport.commandError(diagnostic);
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
    assertEquals(diagnostic, commandError(stdout.toByteArray()));
  }
}
