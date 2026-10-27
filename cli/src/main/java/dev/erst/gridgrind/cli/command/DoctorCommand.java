package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.io.CliRequestReader;
import dev.erst.gridgrind.cli.io.CliResponseWriter;
import dev.erst.gridgrind.cli.render.PathGuidance;
import dev.erst.gridgrind.core.problem.InvalidEncodingException;
import dev.erst.gridgrind.core.problem.InvalidJsonException;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.InvalidRequestShapeException;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.core.problem.RequestDoctorReport;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import dev.erst.gridgrind.engine.api.GridGrindRequestDoctor;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import dev.erst.gridgrind.json.codec.RequestDiagnosticRedactor;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Optional;

/**
 * Handles doctor-request input analysis and transport independently from execution orchestration.
 */
final class DoctorCommand {
  private final CliRequestReader requestReader;
  private final CliResponseWriter responseWriter;
  private final CliDoctorRequestAnalyzer doctorRequestAnalyzer;

  DoctorCommand(
      GridGrindRequestDoctor requestDoctor,
      CliRequestReader requestReader,
      CliResponseWriter responseWriter) {
    this.requestReader = requestReader;
    this.responseWriter = responseWriter;
    this.doctorRequestAnalyzer = new CliDoctorRequestAnalyzer(requestDoctor);
  }

  int run(
      CliCommand.Doctor command,
      InputStream stdin,
      OutputStream stdout,
      OutputStream stderr,
      boolean prettyJson)
      throws IOException {
    if (CliPathArguments.isStandardInputPath(command.requestPath())
        && command.executionRootPath().isEmpty()) {
      return responseWriter.writeCommandError(
          command.responsePath(),
          stdout,
          stderr,
          CommandErrors.invalidArguments(
              "doctor",
              Optional.of("--execution-root"),
              PathGuidance.stdinExecutionRootRequiredMessage()),
          prettyJson);
    }

    RequestDoctorReport report;
    Optional<RequestDiagnosticRedactor> requestRedactor = Optional.empty();
    try {
      byte[] requestBytes = requestReader.readBytes(command.requestPath(), stdin);
      RequestIntake analysis = GridGrindJson.analyzeRequest(requestBytes);
      requestRedactor =
          Optional.of(
              RequestDiagnosticRedactor.forRequestType(
                  dev.erst.gridgrind.core.plan.WorkbookPlan.class));
      report =
          appendAnalysisWarnings(
              doctorRequestAnalyzer.diagnose(
                  command.requestPath(),
                  command.executionRootPath(),
                  command.tempRootPath(),
                  command.grantPath(),
                  command.secretsProviderPath(),
                  analysis,
                  stdin),
              analysis.warnings());
    } catch (InvalidEncodingException
        | InvalidJsonException
        | InvalidRequestShapeException
        | InvalidRequestException exception) {
      return writeFailure(command, stdout, stderr, exception, requestRedactor, prettyJson);
    } catch (IOException exception) {
      return CliRequestReadFailureSupport.write(
          responseWriter,
          "doctor",
          command.requestPath(),
          command.responsePath(),
          stdout,
          stderr,
          exception,
          requestRedactor,
          prettyJson);
    } catch (RuntimeException exception) {
      return writeFailure(command, stdout, stderr, exception, requestRedactor, prettyJson);
    }
    return responseWriter.writeDoctorReport(
        command.responsePath(), stdout, stderr, report, requestRedactor.orElseThrow(), prettyJson);
  }

  private int writeFailure(
      CliCommand.Doctor command,
      OutputStream stdout,
      OutputStream stderr,
      RuntimeException exception,
      Optional<RequestDiagnosticRedactor> requestRedactor,
      boolean prettyJson)
      throws IOException {
    RequestDoctorReport report =
        RequestDoctorReport.invalid(
            Optional.empty(),
            java.util.List.of(),
            GridGrindProblems.fromException(
                exception,
                new ProblemContext.ReadRequest(
                    CliRequestReadFailureSupport.requestInput(command.requestPath()),
                    new ProblemContextRequestSurfaces.JsonLocation.Unavailable())));
    if (requestRedactor.isPresent()) {
      return responseWriter.writeDoctorReport(
          command.responsePath(),
          stdout,
          stderr,
          report,
          requestRedactor.orElseThrow(),
          prettyJson);
    }
    return responseWriter.writeDoctorReport(
        command.responsePath(), stdout, stderr, report, prettyJson);
  }

  private static RequestDoctorReport appendAnalysisWarnings(
      RequestDoctorReport report,
      java.util.List<dev.erst.gridgrind.core.problem.RequestWarning> analysisWarnings) {
    if (analysisWarnings.isEmpty()) {
      return report;
    }
    java.util.List<dev.erst.gridgrind.core.problem.RequestWarning> warnings =
        new java.util.ArrayList<>(report.warnings());
    warnings.addAll(analysisWarnings);
    return report.valid()
        ? RequestDoctorReport.warnings(report.summary().orElseThrow(), warnings)
        : RequestDoctorReport.invalid(report.summary(), warnings, report.problems());
  }
}
