package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.io.CliProgressJsonlSink;
import dev.erst.gridgrind.cli.io.CliRequestReader;
import dev.erst.gridgrind.cli.io.CliResponseReservation;
import dev.erst.gridgrind.cli.io.CliResponseWriter;
import dev.erst.gridgrind.cli.render.PathGuidance;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.InvalidEncodingException;
import dev.erst.gridgrind.core.problem.InvalidJsonException;
import dev.erst.gridgrind.core.problem.InvalidRequestException;
import dev.erst.gridgrind.core.problem.InvalidRequestShapeException;
import dev.erst.gridgrind.core.problem.RequestDoctorReport;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.engine.api.GridGrindRequestDoctor;
import dev.erst.gridgrind.engine.api.GridGrindRequestExecutor;
import dev.erst.gridgrind.engine.api.GridGrindRequestRequirements;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import dev.erst.gridgrind.json.codec.RequestDiagnosticRedactor;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Objects;
import java.util.Optional;

/** Executing and request-doctor CLI flows. */
final class RunCommand {
  private final GridGrindRequestExecutor requestExecutor;
  private final GridGrindRequestDoctor requestDoctor;
  private final CliRequestReader requestReader;
  private final CliResponseWriter responseWriter;

  RunCommand(
      GridGrindRequestExecutor requestExecutor,
      GridGrindRequestDoctor requestDoctor,
      CliRequestReader requestReader,
      CliResponseWriter responseWriter) {
    this.requestExecutor = GridGrindRequestExecutor.requireNonNull(requestExecutor);
    this.requestDoctor = GridGrindRequestDoctor.requireNonNull(requestDoctor);
    this.requestReader = Objects.requireNonNull(requestReader, "requestReader must not be null");
    this.responseWriter = Objects.requireNonNull(responseWriter, "responseWriter must not be null");
  }

  int executeCommand(
      CliCommand.Run command,
      InputStream stdin,
      OutputStream stdout,
      OutputStream stderr,
      boolean prettyJson)
      throws IOException {
    if (CliPathArguments.isStandardInputPath(command.requestPath())
        && command.executionRootPath().isEmpty()) {
      return responseWriter.writeCommandError(
          command.responsePath(), stdout, stderr, stdinExecutionRootFailure("run"), prettyJson);
    }

    WorkbookPlan request;
    RequestIntake analysis;
    Optional<RequestDiagnosticRedactor> requestRedactor = Optional.empty();
    try {
      byte[] requestBytes = requestReader.readBytes(command.requestPath(), stdin);
      analysis = GridGrindJson.analyzeRequest(requestBytes);
      requestRedactor =
          Optional.of(
              RequestDiagnosticRedactor.forRequestType(
                  dev.erst.gridgrind.core.plan.WorkbookPlan.class));
      if (!analysis.isBindable()) {
        RequestDoctorReport staticValidation =
            requestDoctor.diagnose(
                analysis, CliRequestReadFailureSupport.requestInput(command.requestPath()));
        return responseWriter.writeCommandError(
            command.responsePath(),
            stdout,
            stderr,
            CommandErrors.readRequestFailures("run", staticValidation.problems()),
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
          "run",
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
          "run",
          command.requestPath(),
          command.responsePath(),
          stdout,
          stderr,
          exception,
          requestRedactor,
          prettyJson);
    }

    if (CliPathArguments.isStandardInputPath(command.requestPath())
        && GridGrindRequestRequirements.requiresStandardInput(request)) {
      return responseWriter.writeCommandError(
          command.responsePath(),
          stdout,
          stderr,
          CommandErrors.invalidArguments(
              "run", Optional.of("--request"), PathGuidance.standardInputRequiresRequestMessage()),
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
          CommandErrors.readRequestFailures("run", staticValidation.problems()),
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
              new CliExecutionFailureSupport.ExecutionInputs(
                  command.requestPath(),
                  command.executionRootPath(),
                  command.tempRootPath(),
                  command.grantPath(),
                  command.secretsProviderPath(),
                  stdin,
                  new CliProgressJsonlSink(stderr)));
    } catch (IOException exception) {
      return CliRequestReadFailureSupport.write(
          responseWriter,
          "run",
          command.requestPath(),
          command.responsePath(),
          stdout,
          stderr,
          exception,
          requestRedactor,
          prettyJson);
    }

    response = CliResponseWriter.withTransportWarnings(response, analysis);

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
      CliCommand.Run command,
      InputStream stdin,
      OutputStream stdout,
      OutputStream stderr,
      boolean prettyJson,
      WorkbookPlan request,
      RequestIntake analysis,
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
      CliCommand.Run command,
      InputStream stdin,
      OutputStream stderr,
      WorkbookPlan request,
      RequestIntake analysis) {
    try {
      return CliResponseWriter.withTransportWarnings(
          CliExecutionFailureSupport.executeStarted(
              requestExecutor,
              request,
              analysis,
              new CliExecutionFailureSupport.ExecutionInputs(
                  command.requestPath(),
                  command.executionRootPath(),
                  command.tempRootPath(),
                  command.grantPath(),
                  command.secretsProviderPath(),
                  stdin,
                  new CliProgressJsonlSink(stderr))),
          analysis);
    } catch (IOException exception) {
      return CliResponseWriter.withTransportWarnings(
          CliExecutionFailureSupport.failure(request, exception), analysis);
    }
  }

  private static void writeReservationFailureNotice(
      OutputStream stderr, CliResponseReservation.ResponseReservationException exception) {
    try {
      CliResponseWriter.writeTransportNoticeToStderr(
          stderr,
          dev.erst.gridgrind.cli.render.CliTransportNotice.reservationFailure(
              exception.reason(), exception.responsePath().toString()));
    } catch (IOException ignored) {
      // A failed requested transport has no primary payload to reroute before execution begins.
    }
  }

  private static dev.erst.gridgrind.core.result.CommandError stdinExecutionRootFailure(
      String command) {
    return CommandErrors.invalidArguments(
        command, Optional.of("--execution-root"), PathGuidance.stdinExecutionRootRequiredMessage());
  }
}
