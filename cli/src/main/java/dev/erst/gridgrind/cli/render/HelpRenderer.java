package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.cli.command.CliCommand;
import dev.erst.gridgrind.core.discovery.RecipeCatalogEntry;
import dev.erst.gridgrind.core.discovery.RecipeCatalogs;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Pattern;

/** Renders CLI help and its wrapping/alignment from one structured surface. */
public final class HelpRenderer {
  private HelpRenderer() {}

  /** Renders the selected help topic with the packaged product identity. */
  public static String helpText(
      CliCommand.HelpTopic topic,
      String version,
      String description,
      String documentRef,
      String containerTag) {
    Objects.requireNonNull(topic, "topic must not be null");
    Objects.requireNonNull(version, "version must not be null");
    Objects.requireNonNull(description, "description must not be null");
    Objects.requireNonNull(documentRef, "documentRef must not be null");
    Objects.requireNonNull(containerTag, "containerTag must not be null");
    HelpSurfaceReport surface = HelpSurface.CLI_SURFACE;
    String content =
        switch (topic) {
          case OVERVIEW ->
              String.join(
                  "\n\n",
                  renderCommandSection(surface.usage()),
                  renderWorkflows(surface.workflows()),
                  renderSection(surface.execution()),
                  "Use gridgrind help protocol for request constraints, or help guidance for examples.",
                  renderReferences(surface.docs(), documentRef));
          case PROTOCOL ->
              String.join(
                  "\n\n",
                  renderCommandSection(surface.usage()),
                  renderDefinitions(surface.flags()),
                  renderSection(surface.execution()),
                  renderDefinitions(surface.limits()),
                  renderSection(surface.request()),
                  renderDefinitions(surface.fileWorkflow()),
                  renderCoordinateSystems(surface.coordinateSystems()));
          case GUIDANCE ->
              String.join(
                  "\n\n",
                  renderWorkflows(surface.workflows()),
                  renderCommandExample(surface.stdinExample(), containerTag),
                  renderCommandExample(surface.dockerFileExample(), containerTag),
                  renderDiscovery(
                      surface.discovery(), formatRecipes(RecipeCatalogs.catalog().recipes())),
                  renderReferences(surface.docs(), documentRef));
        };
    return ProductInfo.productHeader(version, description) + "\n\n" + content + "\n";
  }

  static String formatRecipes(List<RecipeCatalogEntry> recipes) {
    return recipes.stream()
        .map(HelpRenderer::formatRecipeBlock)
        .collect(java.util.stream.Collectors.joining("\n"));
  }

  static String renderCoordinateSystems(HelpSurfaceReport.CliTableSection section) {
    return renderCoordinateSystems(section, helpTextWidth(System.getenv("COLUMNS")));
  }

  static String renderCoordinateSystems(HelpSurfaceReport.CliTableSection section, int width) {
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

  static String renderSection(HelpSurfaceReport.CliSection section) {
    return renderSection(section, helpTextWidth(System.getenv("COLUMNS")));
  }

  static String renderSection(HelpSurfaceReport.CliSection section, int width) {
    return section.label() + ":\n" + indentLinesWrapped(section.lines(), 2, width);
  }

  static String renderCommandSection(HelpSurfaceReport.CliSection section) {
    return section.label()
        + ":\n"
        + section.lines().stream()
            .map(line -> "  " + line)
            .map(line -> line + "\n")
            .collect(java.util.stream.Collectors.joining())
            .stripTrailing();
  }

  static String renderWorkflows(HelpSurfaceReport.CliWorkflowSection section) {
    return section.label()
        + ":\n"
        + section.entries().stream()
            .map(entry -> "  " + entry.title() + ":\n" + indentLinesWrapped(entry.lines(), 4))
            .collect(java.util.stream.Collectors.joining("\n"));
  }

  static String renderDefinitions(HelpSurfaceReport.CliDefinitionSection section) {
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

  static String renderCommandExample(
      HelpSurfaceReport.CliCommandExample section, String containerTag) {
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

  static String renderDiscovery(HelpSurfaceReport.CliDiscoverySection section, String recipes) {
    return section.label()
        + ":\n"
        + "  Discovery commands:\n"
        + indentLinesWrapped(section.lines(), 4)
        + "\n"
        + "  Unified recipe catalog entries:\n"
        + recipes
        + "\n"
        + "  Print one recipe:\n"
        + "    gridgrind recipe show DASHBOARD --response recipe.json\n"
        + "  Materialize one asset-backed recipe:\n"
        + "    gridgrind recipe materialize CUSTOM_XML --workspace custom-xml\n"
        + renderAdvisoryEntries("Advisory notes", section.guidanceEntries());
  }

  static String renderReferences(
      HelpSurfaceReport.CliReferenceSection section, String documentRef) {
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
    lines.add("    - " + recipe.id() + " (" + recipe.type().name() + ")");
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
      String label, List<HelpSurfaceReport.WorkflowEntry> guidanceEntries) {
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
    return indentLinesWrapped(lines, indentSpaces, helpTextWidth(System.getenv("COLUMNS")));
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

  private static String wrappedDefinition(HelpSurfaceReport.DefinitionEntry entry, int width) {
    String prefix = ("  %-" + width + "s  ").formatted(entry.label() + ":");
    return wrappedText(entry.value(), prefix, " ".repeat(prefix.length()), helpTextWidth());
  }

  private static Optional<LabelAndCommand> splitCommandLabelAndInvocation(String text) {
    return commandLabelAndInvocationFor(text, ": gridgrind ")
        .or(() -> commandLabelAndInvocationFor(text, ": docker "));
  }

  private static int helpTextWidth() {
    return helpTextWidth(System.getenv("COLUMNS"));
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

  private static final int DEFAULT_HELP_TEXT_WIDTH = 88;
  private static final Pattern WHITESPACE_PATTERN = Pattern.compile("\\s+");

  static int helpTextWidth(String columns) {
    if (columns == null || columns.isBlank()) {
      return DEFAULT_HELP_TEXT_WIDTH;
    }
    try {
      return Math.max(72, Math.min(Integer.parseInt(columns.trim()), 120));
    } catch (NumberFormatException ignored) {
      return DEFAULT_HELP_TEXT_WIDTH;
    }
  }

  static String wrappedText(String text, String firstPrefix, String continuationPrefix, int width) {
    Objects.requireNonNull(text, "text must not be null");
    Objects.requireNonNull(firstPrefix, "firstPrefix must not be null");
    Objects.requireNonNull(continuationPrefix, "continuationPrefix must not be null");
    String normalizedText = text.trim();
    StringBuilder builder = new StringBuilder(firstPrefix);
    String currentPrefix = firstPrefix;
    int lineLength = currentPrefix.length();
    for (String token : wrappingTokens(normalizedText)) {
      int separatorWidth = lineLength == currentPrefix.length() ? 0 : 1;
      if (lineLength + separatorWidth + token.length() > width
          && lineLength > currentPrefix.length()) {
        builder.append('\n').append(continuationPrefix).append(token);
        currentPrefix = continuationPrefix;
        lineLength = continuationPrefix.length() + token.length();
      } else {
        if (separatorWidth == 1) {
          builder.append(' ');
          lineLength++;
        }
        builder.append(token);
        lineLength += token.length();
      }
    }
    return builder.toString();
  }

  static List<String> wrappingTokens(String text) {
    List<String> rawTokens =
        WHITESPACE_PATTERN.splitAsStream(text).filter(token -> !token.isBlank()).toList();
    if (rawTokens.isEmpty()) {
      return List.of();
    }
    List<String> tokens = new ArrayList<>();
    int index = 0;
    while (index < rawTokens.size()) {
      String token = rawTokens.get(index);
      if (isRedirectionToken(token) && index + 1 < rawTokens.size()) {
        tokens.add(token + " " + rawTokens.get(index + 1));
        index += 2;
      } else {
        tokens.add(token);
        index++;
      }
    }
    return tokens;
  }

  private static boolean isRedirectionToken(String token) {
    return switch (token) {
      case "<", ">", "1>", "2>", ">>", "1>>", "2>>" -> true;
      default -> false;
    };
  }
}
