package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Parses the immediate non-execution CLI command families. */
final class CliImmediateCommandParser {
  private CliImmediateCommandParser() {}

  static Optional<Result> parse(
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
              new Result(new CliCommand.PrintExampleCatalog(responsePath), index + 1, argument));
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

  private static Result parseHelpCommand(
      int index, CliCommand.HelpTopic topic, Optional<Path> responsePath, String commandToken) {
    return new Result(new CliCommand.Help(topic, responsePath), index + 1, commandToken);
  }

  private static Result parseVersionCommand(int index, Optional<Path> responsePath) {
    return new Result(new CliCommand.Version(responsePath), index + 1, "--version");
  }

  private static Result parseLicenseCommand(int index, Optional<Path> responsePath) {
    return new Result(new CliCommand.License(responsePath), index + 1, "--license");
  }

  private static Result parseRequestTemplateCommand(int index, Optional<Path> responsePath) {
    return new Result(
        new CliCommand.PrintRequestTemplate(responsePath), index + 1, "--print-request-template");
  }

  private static Result parseExampleCommand(String[] args, int index, Optional<Path> responsePath) {
    Optional<String> lookupId = Optional.empty();
    int nextIndex = index + 1;
    boolean keepParsing = true;
    while (nextIndex < args.length && keepParsing) {
      String argument = args[nextIndex];
      if ("--lookup".equals(argument)) {
        if (lookupId.isPresent()) {
          throw new CliArgumentsException("--lookup", "Duplicate argument: --lookup");
        }
        int valueIndex = CliArguments.nextValueIndex(args, nextIndex, "--lookup");
        lookupId =
            Optional.of(
                CliArguments.requireNonBlankValue(
                    "--lookup", args[valueIndex], "example lookup id"));
        nextIndex = valueIndex + 1;
      } else {
        keepParsing = false;
      }
    }
    if (lookupId.isEmpty()) {
      throw new CliArgumentsException(
          "--lookup", "--print-example requires --lookup and one example id value");
    }
    return new Result(
        new CliCommand.PrintExample(lookupId.orElseThrow(), responsePath),
        nextIndex,
        "--print-example");
  }

  private static Result parseTaskCatalogCommand(
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
        int valueIndex = CliArguments.nextValueIndex(args, nextIndex, "--lookup");
        lookupId =
            Optional.of(
                CliArguments.requireNonBlankValue("--lookup", args[valueIndex], "task lookup id"));
        nextIndex = valueIndex + 1;
      } else {
        keepParsing = false;
      }
    }
    return new Result(
        new CliCommand.PrintTaskCatalog(lookupId, responsePath), nextIndex, "--print-task-catalog");
  }

  private static Result parseTaskPlanCommand(
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
        int valueIndex = CliArguments.nextValueIndex(args, nextIndex, "--lookup");
        lookupId =
            Optional.of(
                CliArguments.requireNonBlankValue("--lookup", args[valueIndex], "task lookup id"));
        nextIndex = valueIndex + 1;
      } else {
        keepParsing = false;
      }
    }
    if (lookupId.isEmpty()) {
      throw new CliArgumentsException(
          "--lookup", "--print-task-plan requires --lookup and one task id value");
    }
    return new Result(
        new CliCommand.PrintTaskPlan(lookupId.orElseThrow(), responsePath),
        nextIndex,
        "--print-task-plan");
  }

  private static Result parseTaskKeywordMatchCommand(
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
        int valueIndex = CliArguments.nextValueIndex(args, nextIndex, "--query");
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
    return new Result(
        new CliCommand.PrintTaskKeywordMatch(query.orElseThrow(), responsePath),
        nextIndex,
        "--print-task-keyword-match");
  }

  private static Result parseProtocolCatalogCommand(
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
        int valueIndex = CliArguments.nextValueIndex(args, nextIndex, "--lookup");
        lookupFilter =
            Optional.of(
                CliArguments.requireNonBlankValue(
                    "--lookup", args[valueIndex], "protocol catalog lookup id"));
        nextIndex = valueIndex + 1;
      } else if ("--search".equals(argument)) {
        if (searchQuery.isPresent()) {
          throw new CliArgumentsException("--search", "Duplicate argument: --search");
        }
        int valueIndex = CliArguments.nextValueIndex(args, nextIndex, "--search");
        searchQuery =
            Optional.of(
                CliArguments.requireNonBlankValue("--search", args[valueIndex], "search query"));
        nextIndex = valueIndex + 1;
      } else {
        keepParsing = false;
      }
    }
    return protocolCatalogCommand(lookupFilter, searchQuery, responsePath, nextIndex);
  }

  private static Result protocolCatalogCommand(
      Optional<String> lookupFilter,
      Optional<String> searchQuery,
      Optional<Path> responsePath,
      int nextIndex) {
    if (lookupFilter.isPresent() && searchQuery.isPresent()) {
      throw new CliArgumentsException(
          "--search", "--print-protocol-catalog does not allow both --lookup and --search");
    }
    return new Result(
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

  record Result(CliCommand command, int nextIndex, String commandToken) {
    Result {
      Objects.requireNonNull(command, "command must not be null");
      Objects.requireNonNull(commandToken, "commandToken must not be null");
    }
  }
}
