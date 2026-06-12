package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Parses the execution-oriented CLI grammar once immediate commands are ruled out. */
final class CliExecutionCommandParser {
  private CliExecutionCommandParser() {}

  static CliCommand parse(String[] args, Optional<Path> responsePath) {
    Objects.requireNonNull(args, "args must not be null");
    Objects.requireNonNull(responsePath, "responsePath must not be null");
    ParsedOptions options = new ParsedOptions(responsePath);
    int index = 0;
    while (index < args.length) {
      index = options.consume(args, index, args[index]);
    }
    CliExecutionArgumentValidation.validateTerminalArguments(
        options.requestPath, options.executionRootPath, options.responsePath);
    return options.command();
  }

  static CliArgumentsException unknownArgumentException(String argument) {
    return switch (argument) {
      case "--example" ->
          new CliArgumentsException(
              "--example",
              "--example is no longer part of the CLI grammar; use --print-example --lookup <id>");
      case "--task" ->
          new CliArgumentsException(
              "--task", "--task is no longer part of the CLI grammar; use --lookup instead");
      case "--query" ->
          new CliArgumentsException(
              "--query",
              "--query requires --print-task-keyword-match and one natural-language query value");
      case "--lookup" ->
          new CliArgumentsException(
              "--lookup",
              "--lookup requires --print-example, --print-task-catalog, --print-task-plan,"
                  + " or --print-protocol-catalog and one lookup id value");
      case "--search" ->
          new CliArgumentsException(
              "--search", "--search requires --print-protocol-catalog and one search text value");
      case "--full" ->
          new CliArgumentsException(
              "--full",
              "--full requires --print-protocol-catalog and emits the complete protocol catalog");
      default -> new CliArgumentsException(argument, "Unknown argument: " + argument);
    };
  }

  /** Mutable execution-parser state accumulated while scanning one raw argv array. */
  private static final class ParsedOptions {
    private Optional<Path> requestPath = Optional.empty();
    private Optional<Path> executionRootPath = Optional.empty();
    private Optional<Path> tempRootPath = Optional.empty();
    private final Optional<Path> responsePath;
    private boolean doctorRequest;

    private ParsedOptions(Optional<Path> responsePath) {
      this.responsePath = Objects.requireNonNull(responsePath, "responsePath must not be null");
    }

    private int consume(String[] args, int index, String argument) {
      if ("--doctor-request".equals(argument)) {
        return consumeDoctorRequest(index);
      }
      if ("--request".equals(argument)) {
        return consumeRequestPath(args, index);
      }
      if ("--execution-root".equals(argument)) {
        return consumeExecutionRootPath(args, index);
      }
      if ("--temp-root".equals(argument)) {
        return consumeTempRootPath(args, index);
      }
      if (CliPrimaryCommandSupport.isPrimaryCommandToken(argument)) {
        throw primaryCommandOrderingException(argument);
      }
      throw unknownArgumentException(argument);
    }

    private CliCommand command() {
      return doctorRequest
          ? new CliCommand.DoctorRequest(requestPath, executionRootPath, tempRootPath, responsePath)
          : new CliCommand.Execute(requestPath, executionRootPath, tempRootPath, responsePath);
    }

    private int consumeDoctorRequest(int index) {
      if (doctorRequest) {
        throw new CliArgumentsException("--doctor-request", "Duplicate argument: --doctor-request");
      }
      doctorRequest = true;
      return index + 1;
    }

    private int consumeRequestPath(String[] args, int index) {
      if (requestPath.isPresent()) {
        throw new CliArgumentsException("--request", "Duplicate argument: --request");
      }
      int valueIndex = CliPathArguments.nextValueIndex(args, index, "--request");
      requestPath =
          Optional.of(
              CliPathArguments.requirePathValue(
                  "--request", args[valueIndex], "request path", true));
      return valueIndex + 1;
    }

    private int consumeExecutionRootPath(String[] args, int index) {
      if (executionRootPath.isPresent()) {
        throw new CliArgumentsException("--execution-root", "Duplicate argument: --execution-root");
      }
      int valueIndex = CliPathArguments.nextValueIndex(args, index, "--execution-root");
      executionRootPath =
          Optional.of(
              CliPathArguments.requirePathValue(
                  "--execution-root", args[valueIndex], "execution root path", false));
      return valueIndex + 1;
    }

    private int consumeTempRootPath(String[] args, int index) {
      if (tempRootPath.isPresent()) {
        throw new CliArgumentsException("--temp-root", "Duplicate argument: --temp-root");
      }
      int valueIndex = CliPathArguments.nextValueIndex(args, index, "--temp-root");
      tempRootPath =
          Optional.of(
              CliPathArguments.requirePathValue(
                  "--temp-root", args[valueIndex], "temp root path", false));
      return valueIndex + 1;
    }

    private CliArgumentsException primaryCommandOrderingException(String argument) {
      Objects.requireNonNull(argument, "argument must not be null");
      if (doctorRequest) {
        return new CliArgumentsException(
            argument,
            "Only one primary command may be used per invocation; --doctor-request cannot be combined with "
                + argument);
      }
      return new CliArgumentsException(
          argument,
          argument + " must be the primary command and cannot follow execution arguments");
    }
  }
}
