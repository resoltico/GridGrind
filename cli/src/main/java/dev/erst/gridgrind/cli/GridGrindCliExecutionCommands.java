package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.InvalidJsonException;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.executor.ExecutionInputBindings;
import dev.erst.gridgrind.executor.GridGrindProblems;
import dev.erst.gridgrind.executor.GridGrindRequestDoctor;
import dev.erst.gridgrind.executor.GridGrindRequestExecutor;
import dev.erst.gridgrind.executor.SourceBackedPlanResolver;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PushbackInputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;
import java.util.function.BooleanSupplier;

/** Executing and request-doctor CLI flows. */
final class GridGrindCliExecutionCommands {
  private final GridGrindRequestExecutor requestExecutor;
  private final CliRequestReader requestReader;
  private final CliResponseWriter responseWriter;
  private final CliJournalWriter journalWriter;
  private final BooleanSupplier standardInputIsInteractive;

  GridGrindCliExecutionCommands(
      GridGrindRequestExecutor requestExecutor,
      CliRequestReader requestReader,
      CliResponseWriter responseWriter,
      CliJournalWriter journalWriter,
      BooleanSupplier standardInputIsInteractive) {
    this.requestExecutor = GridGrindRequestExecutor.requireNonNull(requestExecutor);
    this.requestReader = Objects.requireNonNull(requestReader, "requestReader must not be null");
    this.responseWriter = Objects.requireNonNull(responseWriter, "responseWriter must not be null");
    this.journalWriter = Objects.requireNonNull(journalWriter, "journalWriter must not be null");
    this.standardInputIsInteractive =
        Objects.requireNonNull(
            standardInputIsInteractive, "standardInputIsInteractive must not be null");
  }

  Optional<InputStream> standardInputOrNullForImplicitHelp(
      CliCommand.Execute command, String[] args, InputStream stdin, OutputStream stdout)
      throws IOException {
    if (command.requestPath() != null) {
      return Optional.of(stdin);
    }
    if (args.length != 0) {
      return Optional.of(stdin);
    }
    if (standardInputIsInteractive.getAsBoolean()) {
      writeHelp(stdout);
      return Optional.empty();
    }
    PushbackInputStream peekable = new PushbackInputStream(stdin, 1);
    int firstByte = peekable.read();
    if (firstByte < 0) {
      writeHelp(stdout);
      return Optional.empty();
    }
    peekable.unread(firstByte);
    return Optional.of(peekable);
  }

  int executeCommand(
      CliCommand.Execute command, InputStream stdin, OutputStream stdout, OutputStream stderr)
      throws IOException {
    WorkbookPlan request;
    try {
      request = requestReader.read(command.requestPath(), stdin);
    } catch (InvalidJsonException
        | InvalidRequestShapeException
        | InvalidRequestException exception) {
      return responseWriter.write(
          command.responsePath(),
          stdout,
          GridGrindResponses.failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception,
                  new dev.erst.gridgrind.contract.dto.ProblemContext.ReadRequest(
                      requestInput(command.requestPath()),
                      dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation
                          .unavailable()))));
    } catch (IOException exception) {
      return responseWriter.write(
          command.responsePath(),
          stdout,
          GridGrindResponses.failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception,
                  new dev.erst.gridgrind.contract.dto.ProblemContext.ReadRequest(
                      requestInput(command.requestPath()),
                      dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation
                          .unavailable()))));
    }

    GridGrindResponse response;
    if (command.requestPath() == null && SourceBackedPlanResolver.requiresStandardInput(request)) {
      String message = GridGrindContractText.standardInputRequiresRequestMessage();
      response =
          failure(
              GridGrindProblemCode.INVALID_ARGUMENTS,
              message,
              new dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments(
                  dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument.named(
                      "--request")),
              new IllegalArgumentException(message));
      return responseWriter.write(command.responsePath(), stdout, response);
    }

    ExecutionInputBindings bindings =
        CliExecutionBindingsFactory.create(command.requestPath(), request, stdin);
    try {
      response = requestExecutor.execute(request, bindings, journalWriter.sinkFor(request, stderr));
    } catch (Exception exception) {
      response =
          GridGrindResponses.failure(
              GridGrindProtocolVersion.current(),
              GridGrindProblems.fromException(
                  exception,
                  new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteRequest(
                      requestShape(request))));
    }

    return responseWriter.write(command.responsePath(), stdout, response);
  }

  int doctorRequest(CliCommand.DoctorRequest command, InputStream stdin, OutputStream stdout)
      throws IOException {
    RequestDoctorReport report;
    WorkbookPlan request;
    try {
      request = requestReader.read(command.requestPath(), stdin);
    } catch (InvalidJsonException
        | InvalidRequestShapeException
        | InvalidRequestException exception) {
      report =
          RequestDoctorReport.invalid(
              java.util.Optional.empty(),
              java.util.List.of(),
              GridGrindProblems.fromException(
                  exception,
                  new dev.erst.gridgrind.contract.dto.ProblemContext.ReadRequest(
                      requestInput(command.requestPath()),
                      dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation
                          .unavailable())));
      return responseWriter.writeDoctorReport(command.responsePath(), stdout, report);
    } catch (IOException exception) {
      report =
          RequestDoctorReport.invalid(
              java.util.Optional.empty(),
              java.util.List.of(),
              GridGrindProblems.fromException(
                  exception,
                  new dev.erst.gridgrind.contract.dto.ProblemContext.ReadRequest(
                      requestInput(command.requestPath()),
                      dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation
                          .unavailable())));
      return responseWriter.writeDoctorReport(command.responsePath(), stdout, report);
    }

    if (command.requestPath() == null && SourceBackedPlanResolver.requiresStandardInput(request)) {
      report = new GridGrindRequestDoctor().diagnose(request);
      String message = GridGrindContractText.standardInputRequiresRequestMessage();
      report =
          RequestDoctorReport.invalid(
              report.summary(),
              report.warnings(),
              GridGrindProblems.problem(
                  GridGrindProblemCode.INVALID_ARGUMENTS,
                  message,
                  new dev.erst.gridgrind.contract.dto.ProblemContext.ParseArguments(
                      dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument
                          .named("--request")),
                  new IllegalArgumentException(message)));
      return responseWriter.writeDoctorReport(command.responsePath(), stdout, report);
    }

    ExecutionInputBindings bindings =
        command.requestPath() == null
            ? ExecutionInputBindings.processDefault()
            : CliExecutionBindingsFactory.create(command.requestPath(), request, stdin);
    report = new GridGrindRequestDoctor().diagnose(request, bindings);
    return responseWriter.writeDoctorReport(command.responsePath(), stdout, report);
  }

  private void writeHelp(OutputStream stdout) throws IOException {
    byte[] bytes =
        GridGrindCliProductInfo.helpText(GridGrindCliProductInfo.version())
            .getBytes(StandardCharsets.UTF_8);
    responseWriter.writePayload(null, stdout, bytes, 0);
  }

  private static GridGrindResponse.Failure failure(
      GridGrindProblemCode code,
      String message,
      dev.erst.gridgrind.contract.dto.ProblemContext context,
      Throwable cause) {
    return GridGrindResponses.failure(
        GridGrindProtocolVersion.current(),
        GridGrindProblems.problem(code, message, context, cause));
  }

  private dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput requestInput(
      Path path) {
    return path == null
        ? dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput.standardInput()
        : dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput.requestFile(
            path.toAbsolutePath().toString());
  }

  private dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape requestShape(
      WorkbookPlan request) {
    return dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape.known(
        switch (request.source()) {
          case WorkbookPlan.WorkbookSource.New _ -> "NEW";
          case WorkbookPlan.WorkbookSource.ExistingFile _ -> "EXISTING";
        },
        switch (request.persistence()) {
          case WorkbookPlan.WorkbookPersistence.None _ -> "NONE";
          case WorkbookPlan.WorkbookPersistence.OverwriteSource _ -> "OVERWRITE";
          case WorkbookPlan.WorkbookPersistence.SaveAs _ -> "SAVE_AS";
        });
  }
}
