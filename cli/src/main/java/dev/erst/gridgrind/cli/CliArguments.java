package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Parses raw command-line arguments into a typed GridGrind CLI command. */
final class CliArguments {
  private CliArguments() {}

  /** Parses the raw CLI args into the corresponding command model. */
  static CliCommand parse(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    ParsedOptions options = new ParsedOptions();
    String[] remainingArgs = extractGlobalResponse(args, options);

    int index = 0;
    while (index < remainingArgs.length) {
      String argument = remainingArgs[index];
      Optional<ImmediateParseResult> immediate =
          index == 0
              ? parseImmediateCommand(remainingArgs, index, argument, options.responsePath)
              : Optional.empty();
      if (immediate.isPresent()) {
        ImmediateParseResult result = immediate.orElseThrow();
        requireNoTrailingArguments(remainingArgs, result.nextIndex());
        return result.command();
      }
      index = consumeArgument(remainingArgs, index, argument, options);
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

  private static String requireNonBlankValue(String flagName, String value, String description) {
    if (value.isBlank()) {
      throw new CliArgumentsException(flagName, description + " must not be blank");
    }
    return value;
  }

  private static Optional<ImmediateParseResult> parseImmediateCommand(
      String[] args, int index, String argument, Path responsePath) {
    return switch (argument) {
      case "--help", "-h", "help" -> Optional.of(parseHelpCommand(index, responsePath));
      case "--version" -> Optional.of(parseVersionCommand(index, responsePath));
      case "--license" -> Optional.of(parseLicenseCommand(index, responsePath));
      case "--print-request-template" ->
          Optional.of(parseRequestTemplateCommand(index, responsePath));
      case "--print-example" -> {
        int valueIndex = nextValueIndex(args, index, "--print-example");
        yield Optional.of(
            new ImmediateParseResult(
                new CliCommand.PrintExample(
                    requireNonBlankValue("--print-example", args[valueIndex], "example id"),
                    responsePath),
                valueIndex + 1));
      }
      case "--print-task-catalog" ->
          Optional.of(parseTaskCatalogCommand(args, index, responsePath));
      case "--print-task-plan" -> {
        int valueIndex = nextValueIndex(args, index, "--print-task-plan");
        yield Optional.of(
            new ImmediateParseResult(
                new CliCommand.PrintTaskPlan(
                    requireNonBlankValue("--print-task-plan", args[valueIndex], "task id"),
                    responsePath),
                valueIndex + 1));
      }
      case "--print-goal-plan" -> {
        int valueIndex = nextValueIndex(args, index, "--print-goal-plan");
        yield Optional.of(
            new ImmediateParseResult(
                new CliCommand.PrintGoalPlan(
                    requireNonBlankValue("--print-goal-plan", args[valueIndex], "goal"),
                    responsePath),
                valueIndex + 1));
      }
      case "--print-protocol-catalog" ->
          Optional.of(parseProtocolCatalogCommand(args, index, responsePath));
      default -> Optional.empty();
    };
  }

  private static ImmediateParseResult parseHelpCommand(int index, Path responsePath) {
    return new ImmediateParseResult(new CliCommand.Help(responsePath), index + 1);
  }

  private static ImmediateParseResult parseVersionCommand(int index, Path responsePath) {
    return new ImmediateParseResult(new CliCommand.Version(responsePath), index + 1);
  }

  private static ImmediateParseResult parseLicenseCommand(int index, Path responsePath) {
    return new ImmediateParseResult(new CliCommand.License(responsePath), index + 1);
  }

  private static ImmediateParseResult parseRequestTemplateCommand(int index, Path responsePath) {
    return new ImmediateParseResult(new CliCommand.PrintRequestTemplate(responsePath), index + 1);
  }

  private static ImmediateParseResult parseTaskCatalogCommand(
      String[] args, int index, Path responsePath) {
    String taskFilter = null;
    int nextIndex = index + 1;
    boolean keepParsing = true;
    while (nextIndex < args.length && keepParsing) {
      String argument = args[nextIndex];
      if ("--task".equals(argument)) {
        if (taskFilter != null) {
          throw new CliArgumentsException("--task", "Duplicate argument: --task");
        }
        int valueIndex = nextValueIndex(args, nextIndex, "--task");
        taskFilter = requireNonBlankValue("--task", args[valueIndex], "task id");
        nextIndex = valueIndex + 1;
      } else {
        keepParsing = false;
      }
    }
    return new ImmediateParseResult(
        new CliCommand.PrintTaskCatalog(taskFilter, responsePath), nextIndex);
  }

  private static ImmediateParseResult parseProtocolCatalogCommand(
      String[] args, int index, Path responsePath) {
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
        operationFilter =
            requireNonBlankValue("--operation", args[valueIndex], "protocol catalog lookup id");
        nextIndex = valueIndex + 1;
      } else if ("--search".equals(argument)) {
        if (searchQuery != null) {
          throw new CliArgumentsException("--search", "Duplicate argument: --search");
        }
        int valueIndex = nextValueIndex(args, nextIndex, "--search");
        searchQuery = requireNonBlankValue("--search", args[valueIndex], "search query");
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
        protocolCatalogCommand(operationFilter, searchQuery, responsePath), nextIndex);
  }

  private static CliCommand.PrintProtocolCatalog protocolCatalogCommand(
      String operationFilter, String searchQuery, Path responsePath) {
    if (searchQuery != null) {
      return new CliCommand.PrintProtocolCatalogSearch(searchQuery, responsePath);
    }
    if (operationFilter != null) {
      return new CliCommand.PrintProtocolCatalogLookup(operationFilter, responsePath);
    }
    return new CliCommand.PrintProtocolCatalogAll(responsePath);
  }

  private static String[] extractGlobalResponse(String[] args, ParsedOptions options) {
    java.util.List<String> remainingArgs = new java.util.ArrayList<>(args.length);
    int index = 0;
    while (index < args.length) {
      String argument = args[index];
      if (!"--response".equals(argument)) {
        remainingArgs.add(argument);
        index++;
        continue;
      }
      if (options.responsePath != null) {
        throw new CliArgumentsException("--response", "Duplicate argument: --response");
      }
      int valueIndex = nextValueIndex(args, index, "--response");
      options.responsePath = Path.of(args[valueIndex]);
      index = valueIndex + 1;
    }
    return remainingArgs.toArray(String[]::new);
  }

  private static void requireNoTrailingArguments(String[] args, int nextIndex) {
    if (nextIndex == args.length) {
      return;
    }
    String trailingArgument = args[nextIndex];
    throw unknownArgumentException(trailingArgument);
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
      default -> throw unknownArgumentException(argument);
    };
  }

  private static CliArgumentsException unknownArgumentException(String argument) {
    return switch (argument) {
      case "--task" ->
          new CliArgumentsException(
              "--task", "--task requires --print-task-catalog and one task id value");
      case "--operation" ->
          new CliArgumentsException(
              "--operation",
              "--operation requires --print-protocol-catalog and one lookup id value");
      case "--search" ->
          new CliArgumentsException(
              "--search", "--search requires --print-protocol-catalog and one search text value");
      default -> new CliArgumentsException(argument, "Unknown argument: " + argument);
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
}
