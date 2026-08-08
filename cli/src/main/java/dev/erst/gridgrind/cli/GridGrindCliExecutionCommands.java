package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.catalog.GridGrindRequestSurfaceContractText;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.InvalidEncodingException;
import dev.erst.gridgrind.contract.json.InvalidJsonException;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.contract.json.RequestAnalysis;
import dev.erst.gridgrind.contract.json.RequestDiagnosticRedactor;
import dev.erst.gridgrind.engine.api.GridGrindRequestDoctor;
import dev.erst.gridgrind.engine.api.GridGrindRequestExecutor;
import dev.erst.gridgrind.engine.api.GridGrindRequestRequirements;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PushbackInputStream;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;
import java.util.function.BooleanSupplier;

/** Executing and request-doctor CLI flows. */
final class GridGrindCliExecutionCommands {
  private final GridGrindRequestExecutor requestExecutor;
  private final GridGrindRequestDoctor requestDoctor;
  private final CliRequestReader requestReader;
  private final CliResponseWriter responseWriter;
  private final CliJournalWriter journalWriter;
  private final BooleanSupplier standardInputIsInteractive;
  private final CliDoctorRequestAnalyzer doctorRequestAnalyzer;

  GridGrindCliExecutionCommands(
      GridGrindRequestExecutor requestExecutor,
      GridGrindRequestDoctor requestDoctor,
      CliRequestReader requestReader,
      CliResponseWriter responseWriter,
      CliJournalWriter journalWriter,
      BooleanSupplier standardInputIsInteractive) {
    this.requestExecutor = GridGrindRequestExecutor.requireNonNull(requestExecutor);
    this.requestDoctor = GridGrindRequestDoctor.requireNonNull(requestDoctor);
    this.requestReader = Objects.requireNonNull(requestReader, "requestReader must not be null");
    this.responseWriter = Objects.requireNonNull(responseWriter, "responseWriter must not be null");
    this.journalWriter = Objects.requireNonNull(journalWriter, "journalWriter must not be null");
    this.standardInputIsInteractive =
        Objects.requireNonNull(
            standardInputIsInteractive, "standardInputIsInteractive must not be null");
    this.doctorRequestAnalyzer = new CliDoctorRequestAnalyzer(this.requestDoctor);
  }

  Optional<InputStream> standardInputIfPresent(CliCommand.Execute command, InputStream stdin)
      throws IOException {
    return standardInputIfPresent(command.requestPath(), stdin);
  }

  int executeCommand(
      CliCommand.Execute command,
      InputStream stdin,
      OutputStream stdout,
      OutputStream stderr,
      boolean prettyJson)
      throws IOException {
    if (requestArrivesOnStandardInput(command.requestPath())
        && command.executionRootPath().isEmpty()) {
      return responseWriter.writeCommandError(
          command.responsePath(), stdout, stderr, stdinExecutionRootFailure("execute"), prettyJson);
    }

    WorkbookPlan request;
    Optional<RequestDiagnosticRedactor> requestRedactor = Optional.empty();
    try {
      byte[] requestBytes = requestReader.readBytes(command.requestPath(), stdin);
      RequestAnalysis analysis = GridGrindJson.analyzeRequest(requestBytes);
      requestRedactor = Optional.of(analysis.diagnosticRedactor());
      if (!analysis.isBindable()) {
        return responseWriter.writeCommandError(
            command.responsePath(),
            stdout,
            stderr,
            CommandErrors.readRequestFailures(
                "execute",
                CliRequestAnalysisProblems.problems(
                    analysis, CliRequestReadFailureSupport.requestInput(command.requestPath()))),
            requestRedactor.orElseThrow(),
            prettyJson);
      }
      request = analysis.requireCompletePlan();
    } catch (InvalidEncodingException
        | InvalidJsonException
        | InvalidRequestShapeException
        | InvalidRequestException exception) {
      return CliRequestReadFailureSupport.write(
          responseWriter,
          "execute",
          command.requestPath(),
          command.responsePath(),
          stdout,
          stderr,
          exception,
          requestRedactor,
          prettyJson);
    } catch (IOException exception) {
      return CliRequestReadFailureSupport.write(
          responseWriter,
          "execute",
          command.requestPath(),
          command.responsePath(),
          stdout,
          stderr,
          exception,
          requestRedactor,
          prettyJson);
    }

    WorkbookResult response;
    if (requestArrivesOnStandardInput(command.requestPath())
        && GridGrindRequestRequirements.requiresStandardInput(request)) {
      return responseWriter.writeCommandError(
          command.responsePath(),
          stdout,
          stderr,
          CommandErrors.invalidArguments(
              "execute",
              Optional.of("--request"),
              GridGrindRequestSurfaceContractText.standardInputRequiresRequestMessage()),
          requestRedactor.orElseThrow(),
          prettyJson);
    }

    var journalSink = journalWriter.sinkFor(request, stderr);
    try {
      response =
          CliExecutionFailureSupport.executeStarted(
              requestExecutor,
              request,
              command.requestPath(),
              command.executionRootPath(),
              command.tempRootPath(),
              stdin,
              journalSink);
    } catch (IOException exception) {
      return CliRequestReadFailureSupport.write(
          responseWriter,
          "execute",
          command.requestPath(),
          command.responsePath(),
          stdout,
          stderr,
          exception,
          requestRedactor,
          prettyJson);
    }

    return responseWriter.write(
        command.responsePath(),
        stdout,
        stderr,
        response,
        CliExitCodes.forWorkbookResult(response),
        requestRedactor.orElseThrow(),
        prettyJson);
  }

  int doctorRequest(
      CliCommand.DoctorRequest command,
      InputStream stdin,
      OutputStream stdout,
      OutputStream stderr,
      boolean prettyJson)
      throws IOException {
    Optional<InputStream> requestInput = standardInputIfPresent(command.requestPath(), stdin);
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
    if (requestArrivesOnStandardInput(command.requestPath())
        && command.executionRootPath().isEmpty()) {
      return responseWriter.writeCommandError(
          command.responsePath(),
          stdout,
          stderr,
          stdinExecutionRootFailure("doctor-request"),
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
          doctorRequestAnalyzer.diagnose(
              command.requestPath(),
              command.executionRootPath(),
              command.tempRootPath(),
              analysis,
              stdin);
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
    }

    return responseWriter.writeDoctorReport(
        command.responsePath(), stdout, stderr, report, requestRedactor.orElseThrow(), prettyJson);
  }

  private static dev.erst.gridgrind.cli.discovery.CommandError stdinExecutionRootFailure(
      String command) {
    return CommandErrors.invalidArguments(
        command,
        Optional.of("--execution-root"),
        GridGrindRequestSurfaceContractText.stdinExecutionRootRequiredMessage());
  }

  private Optional<InputStream> standardInputIfPresent(
      Optional<Path> requestPath, InputStream stdin) throws IOException {
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    Objects.requireNonNull(stdin, "stdin must not be null");
    if (requestPath.isPresent()) {
      return Optional.of(stdin);
    }
    if (standardInputIsInteractive.getAsBoolean()) {
      return Optional.empty();
    }
    PushbackInputStream peekable = new PushbackInputStream(stdin, 1);
    int firstByte = peekable.read();
    if (firstByte < 0) {
      return Optional.empty();
    }
    peekable.unread(firstByte);
    return Optional.of(peekable);
  }

  private static boolean requestArrivesOnStandardInput(Optional<Path> requestPath) {
    return requestPath.isEmpty() || CliPathArguments.isStandardInputPath(requestPath);
  }
}
