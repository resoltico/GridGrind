package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Parses raw command-line arguments into a typed GridGrind CLI command. */
final class CliArguments {
  private CliArguments() {}

  /** Returns the global {@code --response} path when one was authored in the raw argv array. */
  static Optional<Path> responsePath(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    return authoredPath(args, "--response");
  }

  /** Returns the authored {@code --request} path when one was present in the raw argv array. */
  static Optional<Path> requestPath(String[] args) {
    Objects.requireNonNull(args, "args must not be null");
    return authoredPath(args, "--request");
  }

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
        requireNoTrailingArguments(remainingArgs, result.nextIndex(), result.commandToken());
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
      String[] args, int index, String argument, Optional<Path> responsePath) {
    return switch (argument) {
      case "--help", "-h" ->
          Optional.of(
              parseHelpCommand(index, CliCommand.HelpTopic.OVERVIEW, responsePath, argument));
      case "--help-protocol" ->
          Optional.of(
              parseHelpCommand(index, CliCommand.HelpTopic.PROTOCOL, responsePath, argument));
      case "--help-guidance" ->
          Optional.of(
              parseHelpCommand(index, CliCommand.HelpTopic.GUIDANCE, responsePath, argument));
      case "--version" -> Optional.of(parseVersionCommand(index, responsePath));
      case "--license" -> Optional.of(parseLicenseCommand(index, responsePath));
      case "--print-request-template" ->
          Optional.of(parseRequestTemplateCommand(index, responsePath));
      case "--print-example" -> Optional.of(parseExampleCommand(args, index, responsePath));
      case "--print-example-catalog" ->
          Optional.of(
              new ImmediateParseResult(
                  new CliCommand.PrintExampleCatalog(responsePath), index + 1, argument));
      case "--print-task-catalog" ->
          Optional.of(parseTaskCatalogCommand(args, index, responsePath));
      case "--print-task-plan" -> Optional.of(parseTaskPlanCommand(args, index, responsePath));
      case "--print-task-keyword-match" ->
          Optional.of(parseTaskKeywordMatchCommand(args, index, responsePath));
      case "--print-protocol-catalog" ->
          Optional.of(parseProtocolCatalogCommand(args, index, responsePath));
      default -> Optional.empty();
    };
  }

  private static ImmediateParseResult parseHelpCommand(
      int index, CliCommand.HelpTopic topic, Optional<Path> responsePath, String commandToken) {
    return new ImmediateParseResult(
        new CliCommand.Help(topic, responsePath), index + 1, commandToken);
  }

  private static ImmediateParseResult parseVersionCommand(int index, Optional<Path> responsePath) {
    return new ImmediateParseResult(new CliCommand.Version(responsePath), index + 1, "--version");
  }

  private static ImmediateParseResult parseLicenseCommand(int index, Optional<Path> responsePath) {
    return new ImmediateParseResult(new CliCommand.License(responsePath), index + 1, "--license");
  }

  private static ImmediateParseResult parseRequestTemplateCommand(
      int index, Optional<Path> responsePath) {
    return new ImmediateParseResult(
        new CliCommand.PrintRequestTemplate(responsePath), index + 1, "--print-request-template");
  }

  private static ImmediateParseResult parseExampleCommand(
      String[] args, int index, Optional<Path> responsePath) {
    Optional<String> lookupId = Optional.empty();
    int nextIndex = index + 1;
    boolean keepParsing = true;
    while (nextIndex < args.length && keepParsing) {
      String argument = args[nextIndex];
      if ("--lookup".equals(argument)) {
        if (lookupId.isPresent()) {
          throw new CliArgumentsException("--lookup", "Duplicate argument: --lookup");
        }
        int valueIndex = nextValueIndex(args, nextIndex, "--lookup");
        lookupId =
            Optional.of(requireNonBlankValue("--lookup", args[valueIndex], "example lookup id"));
        nextIndex = valueIndex + 1;
      } else {
        keepParsing = false;
      }
    }
    if (lookupId.isEmpty()) {
      throw new CliArgumentsException(
          "--lookup", "--print-example requires --lookup and one example id value");
    }
    return new ImmediateParseResult(
        new CliCommand.PrintExample(lookupId.orElseThrow(), responsePath),
        nextIndex,
        "--print-example");
  }

  private static ImmediateParseResult parseTaskCatalogCommand(
      String[] args, int index, Optional<Path> responsePath) {
    Optional<String> lookupId = Optional.empty();
    int nextIndex = index + 1;
    boolean keepParsing = true;
    while (nextIndex < args.length && keepParsing) {
      String argument = args[nextIndex];
      if ("--lookup".equals(argument)) {
        if (lookupId.isPresent()) {
          throw new CliArgumentsException("--lookup", "Duplicate argument: --lookup");
        }
        int valueIndex = nextValueIndex(args, nextIndex, "--lookup");
        lookupId =
            Optional.of(requireNonBlankValue("--lookup", args[valueIndex], "task lookup id"));
        nextIndex = valueIndex + 1;
      } else {
        keepParsing = false;
      }
    }
    return new ImmediateParseResult(
        new CliCommand.PrintTaskCatalog(lookupId, responsePath), nextIndex, "--print-task-catalog");
  }

  private static ImmediateParseResult parseTaskPlanCommand(
      String[] args, int index, Optional<Path> responsePath) {
    Optional<String> lookupId = Optional.empty();
    int nextIndex = index + 1;
    boolean keepParsing = true;
    while (nextIndex < args.length && keepParsing) {
      String argument = args[nextIndex];
      if ("--lookup".equals(argument)) {
        if (lookupId.isPresent()) {
          throw new CliArgumentsException("--lookup", "Duplicate argument: --lookup");
        }
        int valueIndex = nextValueIndex(args, nextIndex, "--lookup");
        lookupId =
            Optional.of(requireNonBlankValue("--lookup", args[valueIndex], "task lookup id"));
        nextIndex = valueIndex + 1;
      } else {
        keepParsing = false;
      }
    }
    if (lookupId.isEmpty()) {
      throw new CliArgumentsException(
          "--lookup", "--print-task-plan requires --lookup and one task id value");
    }
    return new ImmediateParseResult(
        new CliCommand.PrintTaskPlan(lookupId.orElseThrow(), responsePath),
        nextIndex,
        "--print-task-plan");
  }

  private static ImmediateParseResult parseTaskKeywordMatchCommand(
      String[] args, int index, Optional<Path> responsePath) {
    Optional<String> query = Optional.empty();
    int nextIndex = index + 1;
    boolean keepParsing = true;
    while (nextIndex < args.length && keepParsing) {
      String argument = args[nextIndex];
      if ("--query".equals(argument)) {
        if (query.isPresent()) {
          throw new CliArgumentsException("--query", "Duplicate argument: --query");
        }
        int valueIndex = nextValueIndex(args, nextIndex, "--query");
        query = Optional.of(args[valueIndex]);
        nextIndex = valueIndex + 1;
      } else {
        keepParsing = false;
      }
    }
    if (query.isEmpty()) {
      throw new CliArgumentsException(
          "--query", "--print-task-keyword-match requires --query and one query value");
    }
    return new ImmediateParseResult(
        new CliCommand.PrintTaskKeywordMatch(query.orElseThrow(), responsePath),
        nextIndex,
        "--print-task-keyword-match");
  }

  private static ImmediateParseResult parseProtocolCatalogCommand(
      String[] args, int index, Optional<Path> responsePath) {
    Optional<String> lookupFilter = Optional.empty();
    Optional<String> searchQuery = Optional.empty();
    int nextIndex = index + 1;
    boolean keepParsing = true;
    while (nextIndex < args.length && keepParsing) {
      String argument = args[nextIndex];
      if ("--lookup".equals(argument)) {
        if (lookupFilter.isPresent()) {
          throw new CliArgumentsException("--lookup", "Duplicate argument: --lookup");
        }
        int valueIndex = nextValueIndex(args, nextIndex, "--lookup");
        lookupFilter =
            Optional.of(
                requireNonBlankValue("--lookup", args[valueIndex], "protocol catalog lookup id"));
        nextIndex = valueIndex + 1;
      } else if ("--search".equals(argument)) {
        if (searchQuery.isPresent()) {
          throw new CliArgumentsException("--search", "Duplicate argument: --search");
        }
        int valueIndex = nextValueIndex(args, nextIndex, "--search");
        searchQuery =
            Optional.of(requireNonBlankValue("--search", args[valueIndex], "search query"));
        nextIndex = valueIndex + 1;
      } else {
        keepParsing = false;
      }
    }
    return protocolCatalogCommand(lookupFilter, searchQuery, responsePath, nextIndex);
  }

  private static ImmediateParseResult protocolCatalogCommand(
      Optional<String> lookupFilter,
      Optional<String> searchQuery,
      Optional<Path> responsePath,
      int nextIndex) {
    if (lookupFilter.isPresent() && searchQuery.isPresent()) {
      throw new CliArgumentsException(
          "--search", "--print-protocol-catalog does not allow both --lookup and --search");
    }
    return new ImmediateParseResult(
        protocolCatalogCommand(lookupFilter, searchQuery, responsePath),
        nextIndex,
        "--print-protocol-catalog");
  }

  private static CliCommand.PrintProtocolCatalog protocolCatalogCommand(
      Optional<String> lookupFilter, Optional<String> searchQuery, Optional<Path> responsePath) {
    if (searchQuery.isPresent()) {
      return new CliCommand.PrintProtocolCatalogSearch(searchQuery.orElseThrow(), responsePath);
    }
    if (lookupFilter.isPresent()) {
      return new CliCommand.PrintProtocolCatalogLookup(lookupFilter.orElseThrow(), responsePath);
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
      if (options.responsePath.isPresent()) {
        throw new CliArgumentsException("--response", "Duplicate argument: --response");
      }
      int valueIndex = nextValueIndex(args, index, "--response");
      options.responsePath = Optional.of(Path.of(args[valueIndex]));
      index = valueIndex + 1;
    }
    return remainingArgs.toArray(String[]::new);
  }

  private static Optional<Path> authoredPath(String[] args, String flagName) {
    Optional<Path> value = Optional.empty();
    int index = 0;
    while (index < args.length) {
      if (!flagName.equals(args[index])) {
        index++;
        continue;
      }
      if (value.isPresent()) {
        throw new CliArgumentsException(flagName, "Duplicate argument: " + flagName);
      }
      int valueIndex = nextValueIndex(args, index, flagName);
      value = Optional.of(Path.of(args[valueIndex]));
      index = valueIndex;
      index++;
    }
    return value;
  }

  private static void requireNoTrailingArguments(
      String[] args, int nextIndex, String commandToken) {
    if (nextIndex == args.length) {
      return;
    }
    String trailingArgument = args[nextIndex];
    if (isPrimaryCommandToken(trailingArgument)) {
      throw new CliArgumentsException(
          trailingArgument,
          "Only one primary command may be used per invocation; "
              + commandToken
              + " cannot be combined with "
              + trailingArgument);
    }
    if ("--request".equals(trailingArgument) || "--doctor-request".equals(trailingArgument)) {
      throw new CliArgumentsException(
          trailingArgument, commandToken + " does not allow " + trailingArgument);
    }
    throw unknownArgumentException(trailingArgument);
  }

  private static boolean isPrimaryCommandToken(String argument) {
    return switch (argument) {
      case "--help",
          "-h",
          "--help-protocol",
          "--help-guidance",
          "--version",
          "--license",
          "--print-request-template",
          "--print-example",
          "--print-example-catalog",
          "--print-task-catalog",
          "--print-task-plan",
          "--print-task-keyword-match",
          "--print-protocol-catalog" ->
          true;
      default -> false;
    };
  }

  private static int consumeArgument(
      String[] args, int index, String argument, ParsedOptions options) {
    return switch (argument) {
      case "--doctor-request" -> {
        if (options.doctorRequest) {
          throw new CliArgumentsException(
              "--doctor-request", "Duplicate argument: --doctor-request");
        }
        options.doctorRequest = true;
        yield index + 1;
      }
      case "--request" -> {
        if (options.requestPath.isPresent()) {
          throw new CliArgumentsException("--request", "Duplicate argument: --request");
        }
        int valueIndex = nextValueIndex(args, index, "--request");
        options.requestPath = Optional.of(Path.of(args[valueIndex]));
        yield valueIndex + 1;
      }
      case "--help",
          "-h",
          "--help-protocol",
          "--help-guidance",
          "--version",
          "--license",
          "--print-request-template",
          "--print-example",
          "--print-example-catalog",
          "--print-task-catalog",
          "--print-task-plan",
          "--print-task-keyword-match",
          "--print-protocol-catalog" ->
          throw primaryCommandOrderingException(argument, options);
      default -> throw unknownArgumentException(argument);
    };
  }

  private static CliArgumentsException primaryCommandOrderingException(
      String argument, ParsedOptions options) {
    Objects.requireNonNull(argument, "argument must not be null");
    Objects.requireNonNull(options, "options must not be null");
    if (options.doctorRequest) {
      return new CliArgumentsException(
          argument,
          "Only one primary command may be used per invocation; --doctor-request cannot be combined with "
              + argument);
    }
    return new CliArgumentsException(
        argument, argument + " must be the primary command and cannot follow execution arguments");
  }

  private static CliArgumentsException unknownArgumentException(String argument) {
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
      default -> new CliArgumentsException(argument, "Unknown argument: " + argument);
    };
  }

  private static void validateTerminalArguments(ParsedOptions options) {
    if (options.requestPath.isPresent()
        && options.responsePath.isPresent()
        && options
            .requestPath
            .orElseThrow()
            .toAbsolutePath()
            .equals(options.responsePath.orElseThrow().toAbsolutePath())) {
      throw new CliArgumentsException(
          "--response", "--request and --response must not point to the same path");
    }
  }

  /** Mutable parser state accumulated while scanning one raw argv array. */
  private static final class ParsedOptions {
    private Optional<Path> requestPath = Optional.empty();
    private Optional<Path> responsePath = Optional.empty();
    private boolean doctorRequest;
  }

  private record ImmediateParseResult(CliCommand command, int nextIndex, String commandToken) {}
}
