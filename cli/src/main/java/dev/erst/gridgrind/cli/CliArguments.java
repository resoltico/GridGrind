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
      Optional<CliImmediateCommandParser.Result> immediate =
          index == 0
              ? CliImmediateCommandParser.parse(
                  remainingArgs, index, argument, options.responsePath)
              : Optional.empty();
      if (immediate.isPresent()) {
        CliImmediateCommandParser.Result result = immediate.orElseThrow();
        requireNoTrailingArguments(remainingArgs, result.nextIndex(), result.commandToken());
        return result.command();
      }
      index = consumeArgument(remainingArgs, index, argument, options);
    }

    validateTerminalArguments(options);
    return options.doctorRequest
        ? new CliCommand.DoctorRequest(
            options.requestPath,
            options.executionRootPath,
            options.tempRootPath,
            options.responsePath)
        : new CliCommand.Execute(
            options.requestPath,
            options.executionRootPath,
            options.tempRootPath,
            options.responsePath);
  }

  static int nextValueIndex(String[] args, int flagIndex, String flagName) {
    int valueIndex = flagIndex + 1;
    if (valueIndex >= args.length) {
      throw new CliArgumentsException(flagName, "Missing value for " + flagName);
    }
    return valueIndex;
  }

  static String requireNonBlankValue(String flagName, String value, String description) {
    if (value.isBlank()) {
      throw new CliArgumentsException(flagName, description + " must not be blank");
    }
    return value;
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
    if ("--request".equals(trailingArgument)
        || "--doctor-request".equals(trailingArgument)
        || "--execution-root".equals(trailingArgument)
        || "--temp-root".equals(trailingArgument)) {
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
      case "--execution-root" -> {
        if (options.executionRootPath.isPresent()) {
          throw new CliArgumentsException(
              "--execution-root", "Duplicate argument: --execution-root");
        }
        int valueIndex = nextValueIndex(args, index, "--execution-root");
        options.executionRootPath = Optional.of(Path.of(args[valueIndex]));
        yield valueIndex + 1;
      }
      case "--temp-root" -> {
        if (options.tempRootPath.isPresent()) {
          throw new CliArgumentsException("--temp-root", "Duplicate argument: --temp-root");
        }
        int valueIndex = nextValueIndex(args, index, "--temp-root");
        options.tempRootPath = Optional.of(Path.of(args[valueIndex]));
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
    if (options.requestPath.isPresent() && options.executionRootPath.isPresent()) {
      throw new CliArgumentsException(
          "--execution-root",
          "--execution-root cannot be combined with --request because the request file directory already owns request-root resolution");
    }
  }

  /** Mutable parser state accumulated while scanning one raw argv array. */
  private static final class ParsedOptions {
    private Optional<Path> requestPath = Optional.empty();
    private Optional<Path> executionRootPath = Optional.empty();
    private Optional<Path> tempRootPath = Optional.empty();
    private Optional<Path> responsePath = Optional.empty();
    private boolean doctorRequest;
  }
}
