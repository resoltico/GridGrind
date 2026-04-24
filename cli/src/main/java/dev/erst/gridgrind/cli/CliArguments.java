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
      ImmediateParseResult immediate =
          index == 0 ? parseImmediateCommand(args, index, argument) : null;
      if (immediate != null) {
        requireNoTrailingArguments(args, immediate.nextIndex());
        return immediate.command();
      }
      index = consumeArgument(args, index, argument, options);
    }

    validateTerminalArguments(options);
    return options.doctorRequest
        ? new CliCommand.DoctorRequest(options.requestPath, options.responsePath)
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
      case "--help", "-h", "help" -> parseHelpCommand(args, index);
      case "--version" -> parseVersionCommand(args, index);
      case "--license" -> parseLicenseCommand(args, index);
      case "--print-request-template" -> parseRequestTemplateCommand(args, index);
      case "--print-example" -> {
        int valueIndex = nextValueIndex(args, index, "--print-example");
        TrailingResponseParseResult response = parseTrailingResponse(args, valueIndex + 1);
        yield new ImmediateParseResult(
            new CliCommand.PrintExample(args[valueIndex], response.responsePath()),
            response.nextIndex());
      }
      case "--print-task-catalog" -> parseTaskCatalogCommand(args, index);
      case "--print-task-plan" -> {
        int valueIndex = nextValueIndex(args, index, "--print-task-plan");
        TrailingResponseParseResult response = parseTrailingResponse(args, valueIndex + 1);
        yield new ImmediateParseResult(
            new CliCommand.PrintTaskPlan(args[valueIndex], response.responsePath()),
            response.nextIndex());
      }
      case "--print-goal-plan" -> {
        int valueIndex = nextValueIndex(args, index, "--print-goal-plan");
        TrailingResponseParseResult response = parseTrailingResponse(args, valueIndex + 1);
        yield new ImmediateParseResult(
            new CliCommand.PrintGoalPlan(args[valueIndex], response.responsePath()),
            response.nextIndex());
      }
      case "--print-protocol-catalog" -> parseProtocolCatalogCommand(args, index);
      default -> null;
    };
  }

  private static ImmediateParseResult parseHelpCommand(String[] args, int index) {
    TrailingResponseParseResult response = parseTrailingResponse(args, index + 1);
    return new ImmediateParseResult(
        new CliCommand.Help(response.responsePath()), response.nextIndex());
  }

  private static ImmediateParseResult parseVersionCommand(String[] args, int index) {
    TrailingResponseParseResult response = parseTrailingResponse(args, index + 1);
    return new ImmediateParseResult(
        new CliCommand.Version(response.responsePath()), response.nextIndex());
  }

  private static ImmediateParseResult parseLicenseCommand(String[] args, int index) {
    TrailingResponseParseResult response = parseTrailingResponse(args, index + 1);
    return new ImmediateParseResult(
        new CliCommand.License(response.responsePath()), response.nextIndex());
  }

  private static ImmediateParseResult parseRequestTemplateCommand(String[] args, int index) {
    TrailingResponseParseResult response = parseTrailingResponse(args, index + 1);
    return new ImmediateParseResult(
        new CliCommand.PrintRequestTemplate(response.responsePath()), response.nextIndex());
  }

  private static ImmediateParseResult parseTaskCatalogCommand(String[] args, int index) {
    String taskFilter = null;
    Path responsePath = null;
    int nextIndex = index + 1;
    boolean keepParsing = true;
    while (nextIndex < args.length && keepParsing) {
      String argument = args[nextIndex];
      if ("--task".equals(argument)) {
        if (taskFilter != null) {
          throw new CliArgumentsException("--task", "Duplicate argument: --task");
        }
        int valueIndex = nextValueIndex(args, nextIndex, "--task");
        taskFilter = args[valueIndex];
        nextIndex = valueIndex + 1;
      } else if ("--response".equals(argument)) {
        if (responsePath != null) {
          throw new CliArgumentsException("--response", "Duplicate argument: --response");
        }
        int valueIndex = nextValueIndex(args, nextIndex, "--response");
        responsePath = Path.of(args[valueIndex]);
        nextIndex = valueIndex + 1;
      } else {
        keepParsing = false;
      }
    }
    return new ImmediateParseResult(
        new CliCommand.PrintTaskCatalog(taskFilter, responsePath), nextIndex);
  }

  private static ImmediateParseResult parseProtocolCatalogCommand(String[] args, int index) {
    String operationFilter = null;
    String searchQuery = null;
    Path responsePath = null;
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
      } else if ("--response".equals(argument)) {
        if (responsePath != null) {
          throw new CliArgumentsException("--response", "Duplicate argument: --response");
        }
        int valueIndex = nextValueIndex(args, nextIndex, "--response");
        responsePath = Path.of(args[valueIndex]);
        nextIndex = valueIndex + 1;
      } else {
        keepParsing = false;
      }
    }
    return protocolCatalogCommand(operationFilter, searchQuery, responsePath, nextIndex);
  }

  private static ImmediateParseResult protocolCatalogCommand(
      String operationFilter, String searchQuery, Path responsePath, int nextIndex) {
    if (operationFilter != null && searchQuery != null) {
      throw new CliArgumentsException(
          "--search", "--print-protocol-catalog does not allow both --operation and --search");
    }
    return new ImmediateParseResult(
        new CliCommand.PrintProtocolCatalog(operationFilter, searchQuery, responsePath), nextIndex);
  }

  private static TrailingResponseParseResult parseTrailingResponse(String[] args, int index) {
    Path responsePath = null;
    int nextIndex = index;
    boolean keepParsing = true;
    while (nextIndex < args.length && keepParsing) {
      String argument = args[nextIndex];
      if ("--response".equals(argument)) {
        if (responsePath != null) {
          throw new CliArgumentsException("--response", "Duplicate argument: --response");
        }
        int valueIndex = nextValueIndex(args, nextIndex, "--response");
        responsePath = Path.of(args[valueIndex]);
        nextIndex = valueIndex + 1;
      } else {
        keepParsing = false;
      }
    }
    return new TrailingResponseParseResult(responsePath, nextIndex);
  }

  private static void requireNoTrailingArguments(String[] args, int nextIndex) {
    if (nextIndex == args.length) {
      return;
    }
    String trailingArgument = args[nextIndex];
    throw new CliArgumentsException(trailingArgument, "Unknown argument: " + trailingArgument);
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
      default -> throw new CliArgumentsException(argument, "Unknown argument: " + argument);
    };
  }

  private static void validateTerminalArguments(ParsedOptions options) {
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

  private record TrailingResponseParseResult(Path responsePath, int nextIndex) {}
}
