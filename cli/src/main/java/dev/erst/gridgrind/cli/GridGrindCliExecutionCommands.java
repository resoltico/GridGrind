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
  private final BooleanSupplier standardInputIsInteractive;
  private final GridGrindCliDoctorCommand doctorCommand;

  GridGrindCliExecutionCommands(
      GridGrindRequestExecutor requestExecutor,
      GridGrindRequestDoctor requestDoctor,
      CliRequestReader requestReader,
      CliResponseWriter responseWriter,
      BooleanSupplier standardInputIsInteractive) {
    this.requestExecutor = GridGrindRequestExecutor.requireNonNull(requestExecutor);
    this.requestDoctor = GridGrindRequestDoctor.requireNonNull(requestDoctor);
    this.requestReader = Objects.requireNonNull(requestReader, "requestReader must not be null");
    this.responseWriter = Objects.requireNonNull(responseWriter, "responseWriter must not be null");
    this.standardInputIsInteractive =
        Objects.requireNonNull(
            standardInputIsInteractive, "standardInputIsInteractive must not be null");
    this.doctorCommand =
        new GridGrindCliDoctorCommand(
            this.requestDoctor,
            this.requestReader,
            this.responseWriter,
            this.standardInputIsInteractive);
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
    RequestAnalysis analysis;
    Optional<RequestDiagnosticRedactor> requestRedactor = Optional.empty();
    try {
      byte[] requestBytes = requestReader.readBytes(command.requestPath(), stdin);
      analysis = GridGrindJson.analyzeRequest(requestBytes);
      requestRedactor = Optional.of(analysis.diagnosticRedactor());
      if (!analysis.isBindable()) {
        RequestDoctorReport staticValidation =
            requestDoctor.diagnose(
                analysis, CliRequestReadFailureSupport.requestInput(command.requestPath()));
        return responseWriter.writeCommandError(
            command.responsePath(),
            stdout,
            stderr,
            CommandErrors.readRequestFailures("execute", staticValidation.problems()),
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
    RequestDoctorReport staticValidation =
        requestDoctor.diagnose(
            analysis, CliRequestReadFailureSupport.requestInput(command.requestPath()));
    if (!staticValidation.problems().isEmpty()) {
      return responseWriter.writeCommandError(
          command.responsePath(),
          stdout,
          stderr,
          CommandErrors.readRequestFailures("execute", staticValidation.problems()),
          requestRedactor.orElseThrow(),
          prettyJson);
    }

    if (command.responsePath().isPresent()) {
      return executeWithReservedResponse(
          command,
          stdin,
          stdout,
          stderr,
          prettyJson,
          request,
          analysis,
          requestRedactor.orElseThrow());
    }

    WorkbookResult response;
    try {
      response =
          CliExecutionFailureSupport.executeStarted(
              requestExecutor,
              request,
              analysis,
              command.requestPath(),
              command.executionRootPath(),
              command.tempRootPath(),
              stdin,
              new CliProgressJsonlSink(stderr));
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

    response = CliResponseAnalysisWarningSupport.append(response, analysis);

    return responseWriter.write(
        command.responsePath(),
        stdout,
        stderr,
        response,
        CliExitCodes.forWorkbookResult(response),
        requestRedactor.orElseThrow(),
        prettyJson);
  }

  private int executeWithReservedResponse(
      CliCommand.Execute command,
      InputStream stdin,
      OutputStream stdout,
      OutputStream stderr,
      boolean prettyJson,
      WorkbookPlan request,
      RequestAnalysis analysis,
      RequestDiagnosticRedactor requestRedactor)
      throws IOException {
    CliResponseReservation reservation;
    try {
      reservation = CliResponseReservation.reserve(command.responsePath().orElseThrow());
    } catch (CliResponseReservation.ResponseReservationException exception) {
      writeReservationFailureNotice(stderr, exception);
      return 1;
    }
    try (reservation) {
      WorkbookResult response = executeReservedRequest(command, stdin, stderr, request, analysis);
      return responseWriter.writeReserved(
          reservation,
          stdout,
          stderr,
          response,
          CliExitCodes.forWorkbookResult(response),
          requestRedactor,
          prettyJson);
    }
  }

  private WorkbookResult executeReservedRequest(
      CliCommand.Execute command,
      InputStream stdin,
      OutputStream stderr,
      WorkbookPlan request,
      RequestAnalysis analysis) {
    try {
      return CliResponseAnalysisWarningSupport.append(
          CliExecutionFailureSupport.executeStarted(
              requestExecutor,
              request,
              analysis,
              command.requestPath(),
              command.executionRootPath(),
              command.tempRootPath(),
              stdin,
              new CliProgressJsonlSink(stderr)),
          analysis);
    } catch (IOException exception) {
      return CliResponseAnalysisWarningSupport.append(
          CliExecutionFailureSupport.failure(request, exception), analysis);
    }
  }

  private static void writeReservationFailureNotice(
      OutputStream stderr, CliResponseReservation.ResponseReservationException exception) {
    try {
      CliResponseTransportSupport.writeTransportNoticeToStderr(
          stderr,
          dev.erst.gridgrind.cli.discovery.CliTransportNotice.reservationFailure(
              exception.reason(), exception.responsePath().toString()));
    } catch (IOException ignored) {
      // A failed requested transport has no primary payload to reroute before execution begins.
    }
  }

  int doctorRequest(
      CliCommand.DoctorRequest command,
      InputStream stdin,
      OutputStream stdout,
      OutputStream stderr,
      boolean prettyJson)
      throws IOException {
    return doctorCommand.run(command, stdin, stdout, stderr, prettyJson);
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
    return CliStandardInputSupport.ifPresent(requestPath, stdin, standardInputIsInteractive);
  }

  private static boolean requestArrivesOnStandardInput(Optional<Path> requestPath) {
    return CliStandardInputSupport.requestArrivesOnStandardInput(requestPath);
  }
}
