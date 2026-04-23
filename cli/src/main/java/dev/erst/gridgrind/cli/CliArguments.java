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
      ImmediateParseResult immediate = parseImmediateCommand(args, index, argument);
      if (immediate != null) {
        if (immediate.command() instanceof CliCommand.PrintProtocolCatalog
            && immediate.nextIndex() != args.length) {
          String trailingArgument = args[immediate.nextIndex()];
          throw new CliArgumentsException(
              trailingArgument, "Unknown argument: " + trailingArgument);
        }
        return immediate.command();
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

  private static ImmediateParseResult parseImmediateCommand(
      String[] args, int index, String argument) {
    return switch (argument) {
      case "--help", "-h" -> new ImmediateParseResult(new CliCommand.Help(), index + 1);
      case "--version" -> new ImmediateParseResult(new CliCommand.Version(), index + 1);
      case "--license" -> new ImmediateParseResult(new CliCommand.License(), index + 1);
      case "--print-request-template" ->
          new ImmediateParseResult(new CliCommand.PrintRequestTemplate(), index + 1);
      case "--print-example" -> {
        int valueIndex = nextValueIndex(args, index, "--print-example");
        yield new ImmediateParseResult(
            new CliCommand.PrintExample(args[valueIndex]), valueIndex + 1);
      }
      case "--print-task-catalog" -> parseTaskCatalogCommand(args, index);
      case "--print-task-plan" -> {
        int valueIndex = nextValueIndex(args, index, "--print-task-plan");
        yield new ImmediateParseResult(
            new CliCommand.PrintTaskPlan(args[valueIndex]), valueIndex + 1);
      }
      case "--print-goal-plan" -> {
        int valueIndex = nextValueIndex(args, index, "--print-goal-plan");
        yield new ImmediateParseResult(
            new CliCommand.PrintGoalPlan(args[valueIndex]), valueIndex + 1);
      }
      case "--print-protocol-catalog" -> parseProtocolCatalogCommand(args, index);
      default -> null;
    };
  }

  private static ImmediateParseResult parseTaskCatalogCommand(String[] args, int index) {
    String taskFilter = null;
    int nextIndex = index + 1;
    if (index + 1 < args.length && "--task".equals(args[index + 1])) {
      int valueIndex = nextValueIndex(args, index + 1, "--task");
      taskFilter = args[valueIndex];
      nextIndex = valueIndex + 1;
    }
    return new ImmediateParseResult(new CliCommand.PrintTaskCatalog(taskFilter), nextIndex);
  }

  private static ImmediateParseResult parseProtocolCatalogCommand(String[] args, int index) {
    String operationFilter = null;
    String searchQuery = null;
    int nextIndex = index + 1;
    boolean keepParsing = true;
    while (nextIndex < args.length && keepParsing) {
      String argument = args[nextIndex];
      if ("--operation".equals(argument)) {
        if (operationFilter != null) {
          throw new CliArgumentsException("--operation", "Duplicate argument: --operation");
        }
        int valueIndex = nextValueIndex(args, nextIndex, "--operation");
        operationFilter = args[valueIndex];
        nextIndex = valueIndex + 1;
      } else if ("--search".equals(argument)) {
        if (searchQuery != null) {
          throw new CliArgumentsException("--search", "Duplicate argument: --search");
        }
        int valueIndex = nextValueIndex(args, nextIndex, "--search");
        searchQuery = args[valueIndex];
        nextIndex = valueIndex + 1;
      } else {
        keepParsing = false;
      }
    }
    return protocolCatalogCommand(operationFilter, searchQuery, nextIndex);
  }

  private static ImmediateParseResult protocolCatalogCommand(
      String operationFilter, String searchQuery, int nextIndex) {
    if (operationFilter != null && searchQuery != null) {
      throw new CliArgumentsException(
          "--search", "--print-protocol-catalog does not allow both --operation and --search");
    }
    return new ImmediateParseResult(
        new CliCommand.PrintProtocolCatalog(operationFilter, searchQuery), nextIndex);
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

  private record ImmediateParseResult(CliCommand command, int nextIndex) {}
}
