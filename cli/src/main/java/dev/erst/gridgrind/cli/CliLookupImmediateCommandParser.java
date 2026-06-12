package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Parses immediate command families that pivot on lookup ids or one natural-language query. */
final class CliLookupImmediateCommandParser {
  private CliLookupImmediateCommandParser() {}

  static CliImmediateCommandParser.Result parseExample(
      String[] args, int index, Optional<Path> responsePath) {
    String lookupId =
        requireLookupValue(
            args,
            index,
            "example lookup id",
            "--print-example requires --lookup and one example id value");
    int nextIndex = consumeLookupSequence(args, index, "example lookup id").nextIndex();
    return new CliImmediateCommandParser.Result(
        new CliCommand.PrintExample(lookupId, responsePath), nextIndex, "--print-example");
  }

  static CliImmediateCommandParser.Result parseTaskCatalog(
      String[] args, int index, Optional<Path> responsePath) {
    LookupSequence lookupSequence = consumeLookupSequence(args, index, "task lookup id");
    return new CliImmediateCommandParser.Result(
        new CliCommand.PrintTaskCatalog(lookupSequence.lookupId(), responsePath),
        lookupSequence.nextIndex(),
        "--print-task-catalog");
  }

  static CliImmediateCommandParser.Result parseTaskPlan(
      String[] args, int index, Optional<Path> responsePath) {
    String lookupId =
        requireLookupValue(
            args,
            index,
            "task lookup id",
            "--print-task-plan requires --lookup and one task id value");
    int nextIndex = consumeLookupSequence(args, index, "task lookup id").nextIndex();
    return new CliImmediateCommandParser.Result(
        new CliCommand.PrintTaskPlan(lookupId, responsePath), nextIndex, "--print-task-plan");
  }

  static CliImmediateCommandParser.Result parseTaskKeywordMatch(
      String[] args, int index, Optional<Path> responsePath) {
    QuerySequence querySequence = consumeQuerySequence(args, index);
    return new CliImmediateCommandParser.Result(
        new CliCommand.PrintTaskKeywordMatch(querySequence.query().orElseThrow(), responsePath),
        querySequence.nextIndex(),
        "--print-task-keyword-match");
  }

  private static String requireLookupValue(
      String[] args, int index, String description, String requirementMessage) {
    LookupSequence lookupSequence = consumeLookupSequence(args, index, description);
    if (lookupSequence.lookupId().isEmpty()) {
      throw new CliArgumentsException("--lookup", requirementMessage);
    }
    return lookupSequence.lookupId().orElseThrow();
  }

  private static LookupSequence consumeLookupSequence(
      String[] args, int index, String description) {
    Objects.requireNonNull(args, "args must not be null");
    Objects.requireNonNull(description, "description must not be null");
    Optional<String> lookupId = Optional.empty();
    int nextIndex = index + 1;
    while (nextIndex < args.length && "--lookup".equals(args[nextIndex])) {
      if (lookupId.isPresent()) {
        throw new CliArgumentsException("--lookup", "Duplicate argument: --lookup");
      }
      int valueIndex = CliPathArguments.nextValueIndex(args, nextIndex, "--lookup");
      lookupId =
          Optional.of(
              CliPathArguments.requireNonBlankValue("--lookup", args[valueIndex], description));
      nextIndex = valueIndex + 1;
    }
    return new LookupSequence(lookupId, nextIndex);
  }

  private static QuerySequence consumeQuerySequence(String[] args, int index) {
    Objects.requireNonNull(args, "args must not be null");
    Optional<String> query = Optional.empty();
    int nextIndex = index + 1;
    while (nextIndex < args.length && "--query".equals(args[nextIndex])) {
      if (query.isPresent()) {
        throw new CliArgumentsException("--query", "Duplicate argument: --query");
      }
      int valueIndex = CliPathArguments.nextValueIndex(args, nextIndex, "--query");
      query = Optional.of(args[valueIndex]);
      nextIndex = valueIndex + 1;
    }
    if (query.isEmpty()) {
      throw new CliArgumentsException(
          "--query", "--print-task-keyword-match requires --query and one query value");
    }
    return new QuerySequence(query, nextIndex);
  }

  record LookupSequence(Optional<String> lookupId, int nextIndex) {
    LookupSequence {
      Objects.requireNonNull(lookupId, "lookupId must not be null");
    }
  }

  record QuerySequence(Optional<String> query, int nextIndex) {
    QuerySequence {
      Objects.requireNonNull(query, "query must not be null");
    }
  }
}
