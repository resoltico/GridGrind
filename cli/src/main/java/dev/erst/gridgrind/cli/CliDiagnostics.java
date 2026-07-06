package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CliDiagnostic;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestShape;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.nio.file.Path;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

/** Shared constructors for machine-readable CLI diagnostics. */
final class CliDiagnostics {
  private CliDiagnostics() {}

  static CliDiagnostic invalidArguments(
      int exitCode,
      String command,
      Optional<String> argument,
      String message,
      List<String> suggestions) {
    ProblemContext.ParseArguments context =
        new ProblemContext.ParseArguments(
            argument.map(CliArgument::named).orElseGet(CliArgument::unknown));
    return diagnostic(
        exitCode,
        command,
        suggestions,
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_ARGUMENTS, message, context));
  }

  static CliDiagnostic responseWriteFailure(
      String command,
      String payloadName,
      Path targetPath,
      java.io.IOException exception,
      Optional<String> stdoutSuggestion) {
    Objects.requireNonNull(targetPath, "targetPath must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    Objects.requireNonNull(stdoutSuggestion, "stdoutSuggestion must not be null");
    ProblemContext.WriteResponse context =
        new ProblemContext.WriteResponse(
            dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.ResponseOutput
                .responseFile(targetPath.toString()));
    String message = CliResponseWriter.responseWriteMessage(exception, targetPath);
    return diagnostic(
        1,
        command,
        stdoutSuggestion.filter(text -> !text.isBlank()).map(List::of).orElse(List.of()),
        GridGrindProblems.problem(
            GridGrindProblemCode.IO_ERROR,
            message,
            context,
            List.of(
                new GridGrindProblemDetail.ProblemCause(
                    GridGrindProblemCode.IO_ERROR, message, context.stage()))));
  }

  static CliDiagnostic readRequestFailure(
      int exitCode, String command, GridGrindProblemDetail.Problem problem) {
    Objects.requireNonNull(problem, "problem must not be null");
    return diagnostic(exitCode, command, suggestionsForReadRequest(problem, command), problem);
  }

  static CliDiagnostic problemDiagnostic(
      int exitCode, String command, GridGrindProblemDetail.Problem problem) {
    Objects.requireNonNull(problem, "problem must not be null");
    return diagnostic(exitCode, command, List.of(), problem);
  }

  static CliDiagnostic unexpectedFailure(String command, Throwable exception) {
    Objects.requireNonNull(command, "command must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    ProblemContext.ExecuteRequest context =
        new ProblemContext.ExecuteRequest(RequestShape.unknown());
    String message =
        Optional.ofNullable(exception.getMessage())
            .filter(text -> !text.isBlank())
            .orElse(GridGrindProblemCode.INTERNAL_ERROR.title());
    return diagnostic(
        1,
        command,
        List.of("gridgrind --help", "gridgrind --help-protocol"),
        GridGrindProblems.problem(
            GridGrindProblemCode.INTERNAL_ERROR,
            message,
            context,
            List.of(
                new GridGrindProblemDetail.ProblemCause(
                    GridGrindProblemCode.INTERNAL_ERROR, message, context.stage()))));
  }

  private static CliDiagnostic diagnostic(
      int exitCode,
      String command,
      List<String> suggestions,
      GridGrindProblemDetail.Problem problem) {
    return new CliDiagnostic(
        GridGrindProtocolVersion.current(),
        exitCode,
        command,
        List.copyOf(Objects.requireNonNull(suggestions, "suggestions must not be null")),
        Objects.requireNonNull(problem, "problem must not be null"),
        Optional.empty());
  }

  private static List<String> suggestionsForReadRequest(
      GridGrindProblemDetail.Problem problem, String command) {
    Set<String> suggestions = new LinkedHashSet<>();
    switch (problem.code()) {
      case INVALID_JSON -> {
        suggestions.add("gridgrind --print-request-template --response request.json");
        suggestions.add("gridgrind --help-protocol");
      }
      case INVALID_REQUEST_SHAPE -> {
        CliSuggestionSupport.protocolCatalogSearchCommandForProblem(problem)
            .ifPresent(suggestions::add);
        if ("doctor-request".equals(command)) {
          suggestions.add("gridgrind --print-request-template --response request.json");
        } else {
          suggestions.add("gridgrind --doctor-request --request request.json");
          suggestions.add("gridgrind --help-protocol");
        }
      }
      case INVALID_REQUEST -> {
        CliSuggestionSupport.protocolCatalogSearchCommandForProblem(problem)
            .ifPresent(suggestions::add);
        if ("doctor-request".equals(command)) {
          suggestions.add("gridgrind --print-request-template --response request.json");
        } else {
          suggestions.add("gridgrind --doctor-request --request request.json");
          suggestions.add("gridgrind --help-protocol");
        }
      }
      case IO_ERROR -> {
        suggestions.add("gridgrind --request request.json");
        suggestions.add("gridgrind --help-guidance");
      }
      default -> {
        suggestions.add("gridgrind --help");
        suggestions.add("gridgrind --help-protocol");
      }
    }
    return List.copyOf(suggestions);
  }
}
