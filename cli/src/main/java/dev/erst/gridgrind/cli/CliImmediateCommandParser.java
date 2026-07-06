package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/** Parses the immediate non-execution CLI command families. */
final class CliImmediateCommandParser {
  private static final Map<String, SimpleCommandFactory> SIMPLE_COMMAND_FACTORIES =
      Map.ofEntries(
          Map.entry(
              "--help",
              (index, responsePath, commandToken) ->
                  parseHelpCommand(
                      index, CliCommand.HelpTopic.OVERVIEW, responsePath, commandToken)),
          Map.entry(
              "help",
              (index, responsePath, commandToken) ->
                  parseHelpCommand(
                      index, CliCommand.HelpTopic.OVERVIEW, responsePath, commandToken)),
          Map.entry(
              "-h",
              (index, responsePath, commandToken) ->
                  parseHelpCommand(
                      index, CliCommand.HelpTopic.OVERVIEW, responsePath, commandToken)),
          Map.entry(
              "--help-protocol",
              (index, responsePath, commandToken) ->
                  parseHelpCommand(
                      index, CliCommand.HelpTopic.PROTOCOL, responsePath, commandToken)),
          Map.entry(
              "help-protocol",
              (index, responsePath, commandToken) ->
                  parseHelpCommand(
                      index, CliCommand.HelpTopic.PROTOCOL, responsePath, commandToken)),
          Map.entry(
              "--help-guidance",
              (index, responsePath, commandToken) ->
                  parseHelpCommand(
                      index, CliCommand.HelpTopic.GUIDANCE, responsePath, commandToken)),
          Map.entry(
              "help-guidance",
              (index, responsePath, commandToken) ->
                  parseHelpCommand(
                      index, CliCommand.HelpTopic.GUIDANCE, responsePath, commandToken)),
          Map.entry(
              "--version",
              (index, responsePath, commandToken) -> parseVersionCommand(index, responsePath)),
          Map.entry(
              "version",
              (index, responsePath, commandToken) -> parseVersionCommand(index, responsePath)),
          Map.entry(
              "--license",
              (index, responsePath, commandToken) -> parseLicenseCommand(index, responsePath)),
          Map.entry(
              "license",
              (index, responsePath, commandToken) -> parseLicenseCommand(index, responsePath)),
          Map.entry(
              "--print-request-template",
              (index, responsePath, commandToken) ->
                  parseRequestTemplateCommand(index, responsePath)));

  private CliImmediateCommandParser() {}

  static Optional<Result> parse(
      String[] args, int index, String argument, Optional<Path> responsePath) {
    Objects.requireNonNull(args, "args must not be null");
    Objects.requireNonNull(argument, "argument must not be null");
    Objects.requireNonNull(responsePath, "responsePath must not be null");
    SimpleCommandFactory simpleCommandFactory = SIMPLE_COMMAND_FACTORIES.get(argument);
    if (simpleCommandFactory != null) {
      return Optional.of(simpleCommandFactory.create(index, responsePath, argument));
    }
    return switch (argument) {
      case "--print-recipe" ->
          Optional.of(CliLookupImmediateCommandParser.parseRecipe(args, index, responsePath));
      case "--print-recipe-catalog" ->
          Optional.of(
              CliLookupImmediateCommandParser.parseRecipeCatalog(args, index, responsePath));
      case "--print-recipe-keyword-match" ->
          Optional.of(
              CliLookupImmediateCommandParser.parseRecipeKeywordMatch(args, index, responsePath));
      case "--print-protocol-catalog" ->
          Optional.of(CliProtocolCatalogCommandParser.parse(args, index, responsePath));
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

  record Result(CliCommand command, int nextIndex, String commandToken) {
    Result {
      Objects.requireNonNull(command, "command must not be null");
      Objects.requireNonNull(commandToken, "commandToken must not be null");
    }
  }

  /** Creates one simple immediate command from the current argument index. */
  @FunctionalInterface
  private interface SimpleCommandFactory {
    /** Returns one parsed simple command and the next index after its consumed arguments. */
    Result create(int index, Optional<Path> responsePath, String commandToken);
  }
}
