package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.InvalidJsonException;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import dev.erst.gridgrind.engine.api.GridGrindRequestDoctor;
import dev.erst.gridgrind.engine.api.GridGrindRequestExecutor;
import dev.erst.gridgrind.engine.api.GridGrindRequestInputs;
import dev.erst.gridgrind.engine.api.GridGrindRequestRequirements;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PushbackInputStream;
import java.nio.file.Path;
import java.util.List;
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
      CliCommand.Execute command, InputStream stdin, OutputStream stdout, OutputStream stderr)
      throws IOException {
    if (requestArrivesOnStandardInput(command.requestPath())
        && command.executionRootPath().isEmpty()) {
      return responseWriter.writeCliFailureReport(
          command.responsePath(), stdout, stderr, stdinExecutionRootFailure("execute"));
    }

    WorkbookPlan request;
    try {
      request = requestReader.read(command.requestPath(), stdin);
    } catch (InvalidJsonException
        | InvalidRequestShapeException
        | InvalidRequestException exception) {
      return writeReadRequestFailure(
          1, "execute", command.requestPath(), command.responsePath(), stdout, stderr, exception);
    } catch (IOException exception) {
      return writeReadRequestFailure(
          1, "execute", command.requestPath(), command.responsePath(), stdout, stderr, exception);
    }

    GridGrindResponse response;
    if (requestArrivesOnStandardInput(command.requestPath())
        && GridGrindRequestRequirements.requiresStandardInput(request)) {
      return responseWriter.writeCliFailureReport(
          command.responsePath(),
          stdout,
          stderr,
          CliFailureReports.invalidArguments(
              2,
              "execute",
              "bind-inputs",
              Optional.of("--request"),
              GridGrindContractText.standardInputRequiresRequestMessage(),
              List.of("gridgrind --request request.json", "gridgrind --help-protocol"),
              Optional.of(
                  "Requests that bind STANDARD_INPUT-authored payloads must arrive by"
                      + " --request so standard input stays available for payload bytes.")));
    }

    GridGrindRequestInputs bindings =
        CliExecutionBindingsFactory.create(
            command.requestPath(),
            command.executionRootPath(),
            command.tempRootPath(),
            request,
            stdin);
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

    return responseWriter.write(command.responsePath(), stdout, stderr, response);
  }

  int doctorRequest(
      CliCommand.DoctorRequest command, InputStream stdin, OutputStream stdout, OutputStream stderr)
      throws IOException {
    Optional<InputStream> requestInput = standardInputIfPresent(command.requestPath(), stdin);
    if (requestInput.isEmpty()) {
      return responseWriter.writeCliFailureReport(
          command.responsePath(),
          stdout,
          stderr,
          CliFailureReports.invalidArguments(
              2,
              "doctor-request",
              "resolve-request",
              Optional.of("--request"),
              "No request JSON was provided. Pass --request <path> or pipe one request document"
                  + " on standard input.",
              List.of(
                  "gridgrind --print-request-template --response request.json",
                  "gridgrind --help",
                  "gridgrind --help-protocol"),
              Optional.of(
                  "Use --doctor-request only after you have one real request document to"
                      + " inspect.")));
    }
    if (requestArrivesOnStandardInput(command.requestPath())
        && command.executionRootPath().isEmpty()) {
      return responseWriter.writeCliFailureReport(
          command.responsePath(), stdout, stderr, stdinExecutionRootFailure("doctor-request"));
    }

    RequestDoctorReport report;
    try {
      byte[] requestBytes =
          requestReader.readBytes(command.requestPath(), requestInput.orElseThrow());
      report =
          doctorRequestAnalyzer.diagnose(
              command.requestPath(),
              command.executionRootPath(),
              command.tempRootPath(),
              requestBytes,
              stdin);
    } catch (IOException exception) {
      return writeReadRequestFailure(
          1,
          "doctor-request",
          command.requestPath(),
          command.responsePath(),
          stdout,
          stderr,
          exception);
    }

    return responseWriter.writeDoctorReport(command.responsePath(), stdout, stderr, report);
  }

  private static dev.erst.gridgrind.cli.discovery.CliFailureReport stdinExecutionRootFailure(
      String command) {
    return CliFailureReports.invalidArguments(
        2,
        command,
        "resolve-request",
        Optional.of("--execution-root"),
        GridGrindContractText.stdinExecutionRootRequiredMessage(),
        List.of(
            "gridgrind --execution-root . < request.json",
            "gridgrind --request request.json",
            "gridgrind --help-protocol"),
        Optional.of(
            "When the request JSON arrives on stdin, pass one explicit --execution-root so"
                + " relative request-owned paths and internal temp files resolve from one"
                + " caller-chosen directory."));
  }

  private dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput requestInput(
      Optional<Path> path) {
    Objects.requireNonNull(path, "path must not be null");
    return requestArrivesOnStandardInput(path)
        ? dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput.standardInput()
        : dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput.requestFile(
            path.orElseThrow().toAbsolutePath().toString());
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

  private int writeReadRequestFailure(
      int exitCode,
      String command,
      Optional<Path> requestPath,
      Optional<Path> responsePath,
      OutputStream stdout,
      OutputStream stderr,
      Throwable exception)
      throws IOException {
    Objects.requireNonNull(command, "command must not be null");
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    Objects.requireNonNull(responsePath, "responsePath must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    return responseWriter.writeRequestFailureReport(
        responsePath,
        stdout,
        stderr,
        CliFailureReports.readRequestFailure(
            exitCode,
            command,
            requestPath.isPresent() ? Optional.of("--request") : Optional.empty(),
            GridGrindProblems.fromException(
                exception,
                new dev.erst.gridgrind.contract.dto.ProblemContext.ReadRequest(
                    requestInput(requestPath),
                    dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.JsonLocation
                        .unavailable())),
            exception));
  }

  private static boolean requestArrivesOnStandardInput(Optional<Path> requestPath) {
    return requestPath.isEmpty() || CliPathArguments.isStandardInputPath(requestPath);
  }
}
