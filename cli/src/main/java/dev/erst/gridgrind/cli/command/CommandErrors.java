package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.HostContext;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.HostArgument;
import dev.erst.gridgrind.core.result.CommandError;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
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
            argument.map(HostArgument::named).orElseGet(HostArgument::unknown));
    return commandError(
        command,
        new GridGrindProblemDetail.Problem(
            GridGrindProblemCode.INVALID_ARGUMENTS,
            GridGrindProblemCode.INVALID_ARGUMENTS.category(),
            GridGrindProblemCode.INVALID_ARGUMENTS.recovery(),
            GridGrindProblemCode.INVALID_ARGUMENTS.title(),
            message,
            ArgumentGuidance.resolution(command, argument),
            context,
            List.of()));
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
    HostContext context = new HostContext("cli");
    String message = GridGrindProblemCode.INTERNAL_ERROR.title();
    return commandError(
        command,
        GridGrindProblems.problem(
            GridGrindProblemCode.INTERNAL_ERROR,
            message,
            context,
            List.of(
                new GridGrindProblemDetail.ProblemCause(
                    GridGrindProblemCode.INTERNAL_ERROR, message, context.type()))));
  }

  private static CommandError commandError(
      String command, List<GridGrindProblemDetail.Problem> problems) {
    return new CommandError(ProtocolVersion.CURRENT, command, problems);
  }

  static CommandError reportFor(String[] args, CliArgumentsException exception) {
    Objects.requireNonNull(args, "args must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    return CommandErrors.invalidArguments(
        CommandName.primary(args),
        Optional.of(exception.argument()),
        Objects.requireNonNullElse(exception.getMessage(), "Invalid command-line arguments"));
  }

  static CommandError reportFor(String[] args, IllegalArgumentException exception) {
    Objects.requireNonNull(args, "args must not be null");
    Objects.requireNonNull(exception, "exception must not be null");
    return CommandErrors.invalidArguments(
        CommandName.primary(args),
        Optional.empty(),
        Objects.requireNonNullElse(exception.getMessage(), "Invalid command-line arguments"));
  }
}
