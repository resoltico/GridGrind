package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.CommandError;
import dev.erst.gridgrind.contract.dto.CliRuntimeContext;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Shared constructors for canonical rejected-command results. */
final class CommandErrors {
  private CommandErrors() {}

  static CommandError invalidArguments(String command, Optional<String> argument, String message) {
    ProblemContext.ParseArguments context =
        new ProblemContext.ParseArguments(
            argument.map(CliArgument::named).orElseGet(CliArgument::unknown));
    return commandError(
        command,
        GridGrindProblemDetail.Problem.of(
            GridGrindProblemCode.INVALID_ARGUMENTS, message, context));
  }

  static CommandError readRequestFailure(String command, GridGrindProblemDetail.Problem problem) {
    Objects.requireNonNull(problem, "problem must not be null");
    return readRequestFailures(command, List.of(problem));
  }

  static CommandError readRequestFailures(
      String command, List<GridGrindProblemDetail.Problem> problems) {
    List<GridGrindProblemDetail.Problem> diagnosticProblems =
        List.copyOf(Objects.requireNonNull(problems, "problems must not be null"));
    if (diagnosticProblems.isEmpty()) {
      throw new IllegalArgumentException("problems must not be empty");
    }
    return commandError(command, diagnosticProblems);
  }

  static CommandError commandError(String command, GridGrindProblemDetail.Problem problem) {
    Objects.requireNonNull(problem, "problem must not be null");
    return commandError(command, List.of(problem));
  }

  static CommandError unexpectedFailure(String command, Throwable exception) {
    Objects.requireNonNull(command, "command must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    CliRuntimeContext context = new CliRuntimeContext();
    String message = GridGrindProblemCode.INTERNAL_ERROR.title();
    return commandError(
        command,
        GridGrindProblems.problem(
            GridGrindProblemCode.INTERNAL_ERROR,
            message,
            context,
            List.of(
                new GridGrindProblemDetail.ProblemCause(
                    GridGrindProblemCode.INTERNAL_ERROR, message, context.stage()))));
  }

  private static CommandError commandError(
      String command, List<GridGrindProblemDetail.Problem> problems) {
    return new CommandError(GridGrindProtocolVersion.current(), command, problems);
  }
}
