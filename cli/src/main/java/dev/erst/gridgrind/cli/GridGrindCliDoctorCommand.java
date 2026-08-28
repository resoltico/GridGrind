package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.InvalidEncodingException;
import dev.erst.gridgrind.contract.json.InvalidJsonException;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.contract.json.RequestAnalysis;
import dev.erst.gridgrind.contract.json.RequestDiagnosticRedactor;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import dev.erst.gridgrind.engine.api.GridGrindRequestDoctor;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Optional;
import java.util.function.BooleanSupplier;

/**
 * Handles doctor-request input analysis and transport independently from execution orchestration.
 */
final class GridGrindCliDoctorCommand {
  private final CliRequestReader requestReader;
  private final CliResponseWriter responseWriter;
  private final BooleanSupplier standardInputIsInteractive;
  private final CliDoctorRequestAnalyzer doctorRequestAnalyzer;

  GridGrindCliDoctorCommand(
      GridGrindRequestDoctor requestDoctor,
      CliRequestReader requestReader,
      CliResponseWriter responseWriter,
      BooleanSupplier standardInputIsInteractive) {
    this.requestReader = requestReader;
    this.responseWriter = responseWriter;
    this.standardInputIsInteractive = standardInputIsInteractive;
    this.doctorRequestAnalyzer = new CliDoctorRequestAnalyzer(requestDoctor);
  }

  int run(
      CliCommand.DoctorRequest command,
      InputStream stdin,
      OutputStream stdout,
      OutputStream stderr,
      boolean prettyJson)
      throws IOException {
    Optional<InputStream> requestInput =
        CliStandardInputSupport.ifPresent(command.requestPath(), stdin, standardInputIsInteractive);
    if (requestInput.isEmpty()) {
      return responseWriter.writeCommandError(
          command.responsePath(),
          stdout,
          stderr,
          CommandErrors.invalidArguments(
              "doctor-request",
              Optional.of("--request"),
              "No request JSON was provided. Pass --request <path> or pipe one request document"
                  + " on standard input."),
          prettyJson);
    }
    if (CliStandardInputSupport.requestArrivesOnStandardInput(command.requestPath())
        && command.executionRootPath().isEmpty()) {
      return responseWriter.writeCommandError(
          command.responsePath(),
          stdout,
          stderr,
          CommandErrors.invalidArguments(
              "doctor-request",
              Optional.of("--execution-root"),
              dev.erst.gridgrind.contract.catalog.GridGrindRequestSurfaceContractText
                  .stdinExecutionRootRequiredMessage()),
          prettyJson);
    }

    RequestDoctorReport report;
    Optional<RequestDiagnosticRedactor> requestRedactor = Optional.empty();
    try {
      byte[] requestBytes =
          requestReader.readBytes(command.requestPath(), requestInput.orElseThrow());
      RequestAnalysis analysis = GridGrindJson.analyzeRequest(requestBytes);
      requestRedactor = Optional.of(analysis.diagnosticRedactor());
      report =
          appendAnalysisWarnings(
              doctorRequestAnalyzer.diagnose(
                  command.requestPath(),
                  command.executionRootPath(),
                  command.tempRootPath(),
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
          "doctor-request",
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
      CliCommand.DoctorRequest command,
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
      java.util.List<dev.erst.gridgrind.contract.dto.RequestWarning> analysisWarnings) {
    if (analysisWarnings.isEmpty()) {
      return report;
    }
    java.util.List<dev.erst.gridgrind.contract.dto.RequestWarning> warnings =
        new java.util.ArrayList<>(report.warnings());
    warnings.addAll(analysisWarnings);
    return report.valid()
        ? RequestDoctorReport.warnings(report.summary().orElseThrow(), warnings)
        : RequestDoctorReport.invalid(report.summary(), warnings, report.problems());
  }
}
