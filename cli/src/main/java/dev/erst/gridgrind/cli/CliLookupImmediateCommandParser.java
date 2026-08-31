package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.Objects;
import java.util.Optional;

/** Parses immediate command families that pivot on lookup ids or one natural-language query. */
final class CliLookupImmediateCommandParser {
  private CliLookupImmediateCommandParser() {}

  static CliImmediateCommandParser.Result parseRecipe(
      String[] args, int index, Optional<Path> responsePath) {
    String lookupId =
        requireLookupValue(
            args,
            index,
            "recipe lookup id",
            "--print-recipe requires --lookup and one recipe id value");
    int nextIndex = consumeLookupSequence(args, index, "recipe lookup id").nextIndex();
    return new CliImmediateCommandParser.Result(
        new CliCommand.PrintRecipe(lookupId, responsePath), nextIndex, "--print-recipe");
  }

  static CliImmediateCommandParser.Result parseMaterializeRecipe(
      String[] args, int index, Optional<Path> responsePath) {
    if (responsePath.isPresent()) {
      throw new CliArgumentsException(
          "--response", "--materialize-recipe does not allow --response");
    }
    MaterializeRecipeArguments materializeArguments =
        consumeMaterializeRecipeArguments(args, index);
    return new CliImmediateCommandParser.Result(
        new CliCommand.MaterializeRecipe(
            materializeArguments.lookupId(), materializeArguments.workspacePath()),
        materializeArguments.nextIndex(),
        "--materialize-recipe");
  }

  private static MaterializeRecipeArguments consumeMaterializeRecipeArguments(
      String[] args, int index) {
    MaterializeRecipeArguments result = MaterializeRecipeArguments.empty(index + 1);
    while (result.nextIndex() < args.length
        && isMaterializeRecipeArgument(args[result.nextIndex()])) {
      result = result.consume(args);
    }
    return result.requireComplete();
  }

  private static boolean isMaterializeRecipeArgument(String argument) {
    return "--lookup".equals(argument) || "--workspace".equals(argument);
  }

  private record MaterializeRecipeArguments(
      Optional<String> optionalLookupId, Optional<Path> optionalWorkspacePath, int nextIndex) {
    private MaterializeRecipeArguments {
      Objects.requireNonNull(optionalLookupId, "optionalLookupId must not be null");
      Objects.requireNonNull(optionalWorkspacePath, "optionalWorkspacePath must not be null");
    }

    static MaterializeRecipeArguments empty(int nextIndex) {
      return new MaterializeRecipeArguments(Optional.empty(), Optional.empty(), nextIndex);
    }

    MaterializeRecipeArguments consume(String[] args) {
      return "--lookup".equals(args[nextIndex]) ? consumeLookup(args) : consumeWorkspace(args);
    }

    private MaterializeRecipeArguments consumeLookup(String[] args) {
      if (optionalLookupId.isPresent()) {
        throw new CliArgumentsException("--lookup", "Duplicate argument: --lookup");
      }
      int valueIndex = CliPathArguments.nextValueIndex(args, nextIndex, "--lookup");
      return new MaterializeRecipeArguments(
          Optional.of(
              CliPathArguments.requireNonBlankValue(
                  "--lookup", args[valueIndex], "recipe lookup id")),
          optionalWorkspacePath,
          valueIndex + 1);
    }

    private MaterializeRecipeArguments consumeWorkspace(String[] args) {
      if (optionalWorkspacePath.isPresent()) {
        throw new CliArgumentsException("--workspace", "Duplicate argument: --workspace");
      }
      int valueIndex = CliPathArguments.nextValueIndex(args, nextIndex, "--workspace");
      return new MaterializeRecipeArguments(
          optionalLookupId,
          Optional.of(
              CliPathArguments.requirePathValue(
                  "--workspace", args[valueIndex], "workspace path", false)),
          valueIndex + 1);
    }

    MaterializeRecipeArguments requireComplete() {
      if (optionalLookupId.isEmpty()) {
        throw new CliArgumentsException(
            "--lookup", "--materialize-recipe requires --lookup and one recipe id value");
      }
      if (optionalWorkspacePath.isEmpty()) {
        throw new CliArgumentsException(
            "--workspace", "--materialize-recipe requires --workspace and one new directory path");
      }
      return this;
    }

    String lookupId() {
      return optionalLookupId.orElseThrow();
    }

    Path workspacePath() {
      return optionalWorkspacePath.orElseThrow();
    }
  }

  static CliImmediateCommandParser.Result parseRecipeCatalog(
      String[] args, int index, Optional<Path> responsePath) {
    LookupSequence lookupSequence = consumeLookupSequence(args, index, "recipe lookup id");
    return new CliImmediateCommandParser.Result(
        new CliCommand.PrintRecipeCatalog(lookupSequence.lookupId(), responsePath),
        lookupSequence.nextIndex(),
        "--print-recipe-catalog");
  }

  static CliImmediateCommandParser.Result parseRecipeKeywordMatch(
      String[] args, int index, Optional<Path> responsePath) {
    QuerySequence querySequence = consumeQuerySequence(args, index);
    return new CliImmediateCommandParser.Result(
        new CliCommand.PrintRecipeKeywordMatch(querySequence.query().orElseThrow(), responsePath),
        querySequence.nextIndex(),
        "--print-recipe-keyword-match");
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
          "--query", "--print-recipe-keyword-match requires --query and one query value");
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
