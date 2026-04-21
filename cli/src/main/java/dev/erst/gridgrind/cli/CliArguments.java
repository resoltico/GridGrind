package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.Objects;

/** Parses raw command-line arguments into a typed GridGrind CLI command. */
final class CliArguments {
  private CliArguments() {}

  /** Parses the raw CLI args into the corresponding command model. */
  static CliCommand parse(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    ParsedOptions options = new ParsedOptions();

    int index = 0;
    while (index < args.length) {
      String argument = args[index];
      CliCommand immediate = parseImmediateCommand(args, index, argument);
      if (immediate != null) {
        return immediate;
      }
      index = consumeArgument(args, index, argument, options);
    }

    validateTerminalArguments(options);
    return options.doctorRequest
        ? new CliCommand.DoctorRequest(options.requestPath)
        : new CliCommand.Execute(options.requestPath, options.responsePath);
  }

  private static int nextValueIndex(String[] args, int flagIndex, String flagName) {
    int valueIndex = flagIndex + 1;
    if (valueIndex >= args.length) {
      throw new CliArgumentsException(flagName, "Missing value for " + flagName);
    }
    return valueIndex;
  }

  private static CliCommand parseImmediateCommand(String[] args, int index, String argument) {
    return switch (argument) {
      case "--help", "-h" -> new CliCommand.Help();
      case "--version" -> new CliCommand.Version();
      case "--license" -> new CliCommand.License();
      case "--print-request-template" -> new CliCommand.PrintRequestTemplate();
      case "--print-example" -> {
        int valueIndex = nextValueIndex(args, index, "--print-example");
        yield new CliCommand.PrintExample(args[valueIndex]);
      }
      case "--print-task-catalog" -> parseTaskCatalogCommand(args, index);
      case "--print-task-plan" -> {
        int valueIndex = nextValueIndex(args, index, "--print-task-plan");
        yield new CliCommand.PrintTaskPlan(args[valueIndex]);
      }
      case "--print-goal-plan" -> {
        int valueIndex = nextValueIndex(args, index, "--print-goal-plan");
        yield new CliCommand.PrintGoalPlan(args[valueIndex]);
      }
      case "--print-protocol-catalog" -> parseProtocolCatalogCommand(args, index);
      default -> null;
    };
  }

  private static CliCommand.PrintTaskCatalog parseTaskCatalogCommand(String[] args, int index) {
    String taskFilter = null;
    if (index + 1 < args.length && "--task".equals(args[index + 1])) {
      int valueIndex = nextValueIndex(args, index + 1, "--task");
      taskFilter = args[valueIndex];
    }
    return new CliCommand.PrintTaskCatalog(taskFilter);
  }

  private static CliCommand.PrintProtocolCatalog parseProtocolCatalogCommand(
      String[] args, int index) {
    String operationFilter = null;
    if (index + 1 < args.length && "--operation".equals(args[index + 1])) {
      int valueIndex = nextValueIndex(args, index + 1, "--operation");
      operationFilter = args[valueIndex];
    }
    return new CliCommand.PrintProtocolCatalog(operationFilter);
  }

  private static int consumeArgument(
      String[] args, int index, String argument, ParsedOptions options) {
    return switch (argument) {
      case "--doctor-request" -> {
        options.doctorRequest = true;
        yield index + 1;
      }
      case "--request" -> {
        if (options.requestPath != null) {
          throw new CliArgumentsException("--request", "Duplicate argument: --request");
        }
        int valueIndex = nextValueIndex(args, index, "--request");
        options.requestPath = Path.of(args[valueIndex]);
        yield valueIndex + 1;
      }
      case "--response" -> {
        if (options.responsePath != null) {
          throw new CliArgumentsException("--response", "Duplicate argument: --response");
        }
        int valueIndex = nextValueIndex(args, index, "--response");
        options.responsePath = Path.of(args[valueIndex]);
        yield valueIndex + 1;
      }
      default -> {
        if (options.doctorRequest) {
          yield index + 1;
        }
        throw new CliArgumentsException(argument, "Unknown argument: " + argument);
      }
    };
  }

  private static void validateTerminalArguments(ParsedOptions options) {
    if (options.doctorRequest && options.responsePath != null) {
      throw new CliArgumentsException(
          "--response", "--response is not supported with --doctor-request");
    }
    if (options.requestPath != null
        && options.responsePath != null
        && options.requestPath.toAbsolutePath().equals(options.responsePath.toAbsolutePath())) {
      throw new CliArgumentsException(
          "--response", "--request and --response must not point to the same path");
    }
  }

  /** Mutable parser state accumulated while scanning one raw argv array. */
  private static final class ParsedOptions {
    private Path requestPath;
    private Path responsePath;
    private boolean doctorRequest;
  }
}
