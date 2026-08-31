package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.RecipeCatalogEntry;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Pattern;

/** Shared formatting helpers for CLI-owned public help surfaces. */
final class GridGrindCliHelpRenderSupport {
  private GridGrindCliHelpRenderSupport() {}

  static String formatRecipes(List<RecipeCatalogEntry> recipes) {
    return recipes.stream()
        .map(GridGrindCliHelpRenderSupport::formatRecipeBlock)
        .collect(java.util.stream.Collectors.joining("\n"));
  }

  static String renderCoordinateSystems(CliSurface.CliTableSection section) {
    return renderCoordinateSystems(
        section, GridGrindCliWrappingSupport.helpTextWidth(System.getenv("COLUMNS")));
  }

  static String renderCoordinateSystems(CliSurface.CliTableSection section, int width) {
    int leftWidth =
        Math.max(
            section.leftHeader().length(),
            section.entries().stream().mapToInt(entry -> entry.pattern().length()).max().orElse(0));
    if (leftWidth + section.rightHeader().length() + 6 > width) {
      return section.label()
          + ":\n"
          + section.entries().stream()
              .map(
                  entry ->
                      wrappedText(entry.pattern() + ": " + entry.convention(), "  ", "    ", width))
              .map(line -> line + "\n")
              .collect(java.util.stream.Collectors.joining())
              .stripTrailing();
    }
    String separator =
        "  " + "-".repeat(leftWidth) + " " + "-".repeat(section.rightHeader().length());
    return section.label()
        + ":\n"
        + ("  %-" + leftWidth + "s %s\n").formatted(section.leftHeader(), section.rightHeader())
        + separator
        + "\n"
        + section.entries().stream()
            .map(
                entry ->
                    ("  %-" + leftWidth + "s %s").formatted(entry.pattern(), entry.convention()))
            .map(line -> line + "\n")
            .collect(java.util.stream.Collectors.joining())
            .stripTrailing();
  }

  static String renderSection(CliSurface.CliSection section) {
    return renderSection(
        section, GridGrindCliWrappingSupport.helpTextWidth(System.getenv("COLUMNS")));
  }

  static String renderSection(CliSurface.CliSection section, int width) {
    return section.label() + ":\n" + indentLinesWrapped(section.lines(), 2, width);
  }

  static String renderCommandSection(CliSurface.CliSection section) {
    return section.label()
        + ":\n"
        + section.lines().stream()
            .map(line -> "  " + line)
            .map(line -> line + "\n")
            .collect(java.util.stream.Collectors.joining())
            .stripTrailing();
  }

  static String renderWorkflows(CliSurface.CliWorkflowSection section) {
    return section.label()
        + ":\n"
        + section.entries().stream()
            .map(entry -> "  " + entry.title() + ":\n" + indentLinesWrapped(entry.lines(), 4))
            .collect(java.util.stream.Collectors.joining("\n"));
  }

  static String renderDefinitions(CliSurface.CliDefinitionSection section) {
    if (section.entries().isEmpty()) {
      return section.label() + ":";
    }
    int width =
        section.entries().stream().mapToInt(entry -> entry.label().length() + 1).max().orElse(0);
    return section.label()
        + ":\n"
        + section.entries().stream()
            .map(entry -> wrappedDefinition(entry, width))
            .map(block -> block + "\n")
            .collect(java.util.stream.Collectors.joining())
            .stripTrailing();
  }

  static String renderCommandExample(CliSurface.CliCommandExample section, String containerTag) {
    String commands =
        section.commandLines().stream()
            .map(line -> replacePlaceholders(line, containerTag))
            .map(line -> "  " + line)
            .map(line -> line + "\n")
            .collect(java.util.stream.Collectors.joining())
            .stripTrailing();
    if (section.description().isEmpty()) {
      return section.label() + ":\n" + commands;
    }
    return section.label()
        + ":\n"
        + commands
        + "\n\n"
        + indentLinesWrapped(
            List.of(replacePlaceholders(section.description().orElseThrow(), containerTag)), 2);
  }

  static String renderDiscovery(CliSurface.CliDiscoverySection section, String recipes) {
    return section.label()
        + ":\n"
        + "  Discovery commands:\n"
        + indentLinesWrapped(section.lines(), 4)
        + "\n"
        + "  Unified recipe catalog entries:\n"
        + recipes
        + "\n"
        + "  Print one recipe:\n"
        + "    gridgrind --print-recipe --lookup DASHBOARD --response recipe.json\n"
        + "  Materialize one asset-backed recipe:\n"
        + "    gridgrind --materialize-recipe --lookup CUSTOM_XML --workspace custom-xml\n"
        + renderAdvisoryEntries("Advisory notes", section.guidanceEntries());
  }

  static String renderReferences(CliSurface.CliReferenceSection section, String documentRef) {
    if (section.entries().isEmpty()) {
      return section.label() + ":";
    }
    return section.label()
        + ":\n"
        + section.entries().stream()
            .map(
                entry ->
                    "  "
                        + entry.label()
                        + ": "
                        + documentRef
                        + "/"
                        + entry.relativePath()
                        + "\n"
                        + wrappedText(entry.description(), "    ", "    ", helpTextWidth()))
            .map(block -> block + "\n")
            .collect(java.util.stream.Collectors.joining())
            .stripTrailing();
  }

  private static String formatRecipeBlock(RecipeCatalogEntry recipe) {
    List<String> lines = new ArrayList<>();
    lines.add("    - " + recipe.id() + " (" + recipe.view().name() + ")");
    lines.add(
        wrappedText(
            "requestFileName: " + recipe.requestFileName(), "      ", "        ", helpTextWidth()));
    lines.add(
        wrappedText(
            "advisory: " + recipe.advisory().name(), "      ", "        ", helpTextWidth()));
    lines.add(wrappedText("summary: " + recipe.summary(), "      ", "        ", helpTextWidth()));
    if (!recipe.requiredWorkspacePaths().isEmpty()) {
      lines.add(
          wrappedText(
              "requiredWorkspacePaths: " + String.join(", ", recipe.requiredWorkspacePaths()),
              "      ",
              "        ",
              helpTextWidth()));
    }
    return String.join("\n", lines);
  }

  private static String renderAdvisoryEntries(
      String label, List<CliSurface.WorkflowEntry> guidanceEntries) {
    return "  "
        + label
        + ":\n"
        + guidanceEntries.stream()
            .map(
                entry ->
                    wrappedText(entry.title() + ".", "    - ", "      ", helpTextWidth())
                        + "\n"
                        + indentLinesWrapped(entry.lines(), 6))
            .collect(java.util.stream.Collectors.joining("\n"));
  }

  private static String indentLinesWrapped(List<String> lines, int indentSpaces) {
    return indentLinesWrapped(
        lines, indentSpaces, GridGrindCliWrappingSupport.helpTextWidth(System.getenv("COLUMNS")));
  }

  private static String indentLinesWrapped(List<String> lines, int indentSpaces, int width) {
    String indent = " ".repeat(indentSpaces);
    return lines.stream()
        .map(line -> wrappedIndentedLine(line, indent, width))
        .map(block -> block + "\n")
        .collect(java.util.stream.Collectors.joining())
        .stripTrailing();
  }

  static String wrappedIndentedLine(String line, String indent, int width) {
    Optional<LabelAndCommand> labelAndCommand = splitCommandLabelAndInvocation(line.trim());
    if (labelAndCommand.isPresent()) {
      return wrappedCommandExampleLine(indent, labelAndCommand.orElseThrow(), width);
    }
    java.util.regex.Matcher orderedMatcher =
        Pattern.compile("^(\\d+\\.)\\s+(.*)$").matcher(line.trim());
    if (orderedMatcher.matches()) {
      String firstPrefix = indent + orderedMatcher.group(1) + " ";
      String continuationPrefix = " ".repeat(firstPrefix.length());
      return wrappedText(orderedMatcher.group(2), firstPrefix, continuationPrefix, width);
    }
    java.util.regex.Matcher bulletMatcher =
        Pattern.compile("^([-*])\\s+(.*)$").matcher(line.trim());
    if (bulletMatcher.matches()) {
      String firstPrefix = indent + bulletMatcher.group(1) + " ";
      String continuationPrefix = " ".repeat(firstPrefix.length());
      return wrappedText(bulletMatcher.group(2), firstPrefix, continuationPrefix, width);
    }
    return wrappedText(line, indent, indent, width);
  }

  private static String wrappedCommandExampleLine(
      String indent, LabelAndCommand labelAndCommand, int width) {
    String label = labelAndCommand.label();
    String command = labelAndCommand.command();
    java.util.regex.Matcher orderedMatcher = Pattern.compile("^(\\d+\\.)\\s+(.*)$").matcher(label);
    if (orderedMatcher.matches()) {
      String firstPrefix = indent + orderedMatcher.group(1) + " ";
      String continuationPrefix = " ".repeat(firstPrefix.length());
      return wrappedText(orderedMatcher.group(2), firstPrefix, continuationPrefix, width)
          + "\n"
          + wrappedText(command, continuationPrefix, continuationPrefix, width);
    }
    java.util.regex.Matcher bulletMatcher = Pattern.compile("^([-*])\\s+(.*)$").matcher(label);
    if (bulletMatcher.matches()) {
      String firstPrefix = indent + bulletMatcher.group(1) + " ";
      String continuationPrefix = " ".repeat(firstPrefix.length());
      return wrappedText(bulletMatcher.group(2), firstPrefix, continuationPrefix, width)
          + "\n"
          + wrappedText(command, continuationPrefix, continuationPrefix, width);
    }
    String continuationPrefix = indent + "  ";
    return wrappedText(label, indent, continuationPrefix, width)
        + "\n"
        + wrappedText(command, continuationPrefix, continuationPrefix, width);
  }

  private static String wrappedDefinition(CliSurface.DefinitionEntry entry, int width) {
    String prefix = ("  %-" + width + "s  ").formatted(entry.label() + ":");
    return wrappedText(entry.value(), prefix, " ".repeat(prefix.length()), helpTextWidth());
  }

  private static String wrappedText(
      String text, String firstPrefix, String continuationPrefix, int width) {
    return GridGrindCliWrappingSupport.wrappedText(text, firstPrefix, continuationPrefix, width);
  }

  private static Optional<LabelAndCommand> splitCommandLabelAndInvocation(String text) {
    return commandLabelAndInvocationFor(text, ": gridgrind ")
        .or(() -> commandLabelAndInvocationFor(text, ": docker "));
  }

  private static int helpTextWidth() {
    return GridGrindCliWrappingSupport.helpTextWidth(System.getenv("COLUMNS"));
  }

  private static String replacePlaceholders(String value, String containerTag) {
    return value.replace("{{CONTAINER_TAG}}", containerTag);
  }

  private static Optional<LabelAndCommand> commandLabelAndInvocationFor(
      String text, String marker) {
    int markerIndex = text.indexOf(marker);
    if (markerIndex < 0) {
      return Optional.empty();
    }
    return Optional.of(
        new LabelAndCommand(
            text.substring(0, markerIndex + 1).trim(), text.substring(markerIndex + 2).trim()));
  }

  private record LabelAndCommand(String label, String command) {
    private LabelAndCommand {
      Objects.requireNonNull(label, "label must not be null");
      Objects.requireNonNull(command, "command must not be null");
    }
  }
}
