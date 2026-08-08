package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.catalog.GridGrindRequestSurfaceContractText;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponses;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.InvalidEncodingException;
import dev.erst.gridgrind.contract.json.InvalidJsonException;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.contract.json.RequestAnalysis;
import dev.erst.gridgrind.contract.json.RequestDiagnosticRedactor;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import dev.erst.gridgrind.engine.api.GridGrindRequestDoctor;
import dev.erst.gridgrind.engine.api.GridGrindRequestExecutor;
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
      CliCommand.Execute command,
      InputStream stdin,
      OutputStream stdout,
      OutputStream stderr,
      boolean prettyJson)
      throws IOException {
    if (requestArrivesOnStandardInput(command.requestPath())
        && command.executionRootPath().isEmpty()) {
      return responseWriter.writeCliDiagnostic(
          command.responsePath(), stdout, stderr, stdinExecutionRootFailure("execute"), prettyJson);
    }

    WorkbookPlan request;
    Optional<RequestDiagnosticRedactor> requestRedactor = Optional.empty();
    try {
      byte[] requestBytes = requestReader.readBytes(command.requestPath(), stdin);
      RequestAnalysis analysis = GridGrindJson.analyzeRequest(requestBytes);
      requestRedactor = Optional.of(analysis.diagnosticRedactor());
      if (!analysis.isBindable()) {
        return responseWriter.writeRequestDiagnostic(
            command.responsePath(),
            stdout,
            stderr,
            CliDiagnostics.readRequestFailures(
                1,
                "execute",
                CliRequestAnalysisProblems.problems(analysis, requestInput(command.requestPath()))),
            requestRedactor.orElseThrow(),
            prettyJson);
      }
      request = analysis.requireCompletePlan();
    } catch (InvalidEncodingException
        | InvalidJsonException
        | InvalidRequestShapeException
        | InvalidRequestException exception) {
      return writeReadRequestFailure(
          1,
          "execute",
          command.requestPath(),
          command.responsePath(),
          stdout,
          stderr,
          exception,
          requestRedactor,
          prettyJson);
    } catch (IOException exception) {
      return writeReadRequestFailure(
          1,
          "execute",
          command.requestPath(),
          command.responsePath(),
          stdout,
          stderr,
          exception,
          requestRedactor,
          prettyJson);
    }

    GridGrindResponse response;
    if (requestArrivesOnStandardInput(command.requestPath())
        && GridGrindRequestRequirements.requiresStandardInput(request)) {
      return responseWriter.writeRequestDiagnostic(
          command.responsePath(),
          stdout,
          stderr,
          CliDiagnostics.invalidArguments(
              2,
              "execute",
              Optional.of("--request"),
              GridGrindRequestSurfaceContractText.standardInputRequiresRequestMessage(),
              List.of("gridgrind --request request.json", "gridgrind --help-protocol")),
          requestRedactor.orElseThrow(),
          prettyJson);
    }

    try {
      try (CliExecutionBindingsFactory.ManagedRequestInputs bindings =
          CliExecutionBindingsFactory.create(
              command.requestPath(),
              command.executionRootPath(),
              command.tempRootPath(),
              request,
              stdin)) {
        try {
          response =
              requestExecutor.execute(
                  request, bindings.inputs(), journalWriter.sinkFor(request, stderr));
        } catch (Exception exception) {
          response =
              GridGrindResponses.failure(
                  dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(),
                  GridGrindProblems.fromException(
                      exception,
                      new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteRequest(
                          requestShape(request))));
        }
      }
    } catch (IOException exception) {
      return writeReadRequestFailure(
          1,
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
        CliResponseTransportSupport.exitCodeFor(response),
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
      return responseWriter.writeCliDiagnostic(
          command.responsePath(),
          stdout,
          stderr,
          CliDiagnostics.invalidArguments(
              2,
              "doctor-request",
              Optional.of("--request"),
              "No request JSON was provided. Pass --request <path> or pipe one request document"
                  + " on standard input.",
              List.of(
                  "gridgrind --print-request-template --response request.json",
                  "gridgrind --help",
                  "gridgrind --help-protocol")),
          prettyJson);
    }
    if (requestArrivesOnStandardInput(command.requestPath())
        && command.executionRootPath().isEmpty()) {
      return responseWriter.writeCliDiagnostic(
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
      return writeReadRequestFailure(
          1,
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

  private static dev.erst.gridgrind.cli.discovery.CliDiagnostic stdinExecutionRootFailure(
      String command) {
    return CliDiagnostics.invalidArguments(
        2,
        command,
        Optional.of("--execution-root"),
        GridGrindRequestSurfaceContractText.stdinExecutionRootRequiredMessage(),
        List.of(
            "gridgrind --execution-root . < request.json",
            "gridgrind --request request.json",
            "gridgrind --help-protocol"));
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
          case WorkbookPlan.WorkbookPersistence.Overwrite _ -> "OVERWRITE";
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
      Throwable exception,
      Optional<RequestDiagnosticRedactor> requestRedactor,
      boolean prettyJson)
      throws IOException {
    Objects.requireNonNull(command, "command must not be null");
    Objects.requireNonNull(requestPath, "requestPath must not be null");
    Objects.requireNonNull(responsePath, "responsePath must not be null");
    Objects.requireNonNull(stdout, "stdout must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    Objects.requireNonNull(requestRedactor, "requestRedactor must not be null");
    dev.erst.gridgrind.contract.dto.GridGrindProblemDetail.Problem problem =
        GridGrindProblems.fromException(
            exception,
            new dev.erst.gridgrind.contract.dto.ProblemContext.ReadRequest(
                requestInput(requestPath),
                CliRequestFailureLocationSupport.locationFor(exception)));
    var diagnostic = CliDiagnostics.readRequestFailure(exitCode, command, problem);
    if (requestRedactor.isPresent()) {
      return responseWriter.writeRequestDiagnostic(
          responsePath, stdout, stderr, diagnostic, requestRedactor.orElseThrow(), prettyJson);
    }
    return responseWriter.writeRequestDiagnostic(
        responsePath, stdout, stderr, diagnostic, prettyJson);
  }

  private static boolean requestArrivesOnStandardInput(Optional<Path> requestPath) {
    return requestPath.isEmpty() || CliPathArguments.isStandardInputPath(requestPath);
  }
}
