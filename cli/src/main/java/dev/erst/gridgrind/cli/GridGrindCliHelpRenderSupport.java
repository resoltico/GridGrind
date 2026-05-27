package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.ShippedExampleEntry;
import dev.erst.gridgrind.cli.discovery.TaskEntry;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.regex.Pattern;

/** Shared formatting helpers for CLI-owned public help surfaces. */
final class GridGrindCliHelpRenderSupport {
  private static final int HELP_TEXT_WIDTH = 96;
  private static final Pattern WHITESPACE_PATTERN = Pattern.compile("\\s+");

  private GridGrindCliHelpRenderSupport() {}

  static String formatExamples(List<ShippedExampleEntry> shippedExamples) {
    int idWidth =
        shippedExamples.stream().mapToInt(example -> example.id().length()).max().orElse(0);
    int pathWidth =
        shippedExamples.stream()
            .mapToInt(example -> example.suggestedRequestPath().length())
            .max()
            .orElse(0);
    int workspaceModeWidth =
        shippedExamples.stream()
            .mapToInt(example -> example.workspaceMode().name().length())
            .max()
            .orElse(0);
    int labelIdWidth = Math.max(idWidth, "Example Id".length());
    int labelPathWidth = Math.max(pathWidth, "Suggested Request Path".length());
    int labelWorkspaceModeWidth = Math.max(workspaceModeWidth, "Workspace Mode".length());
    String header =
        ("    %-"
                + labelIdWidth
                + "s  %-"
                + labelPathWidth
                + "s  %-"
                + labelWorkspaceModeWidth
                + "s  %s")
            .formatted("Example Id", "Suggested Request Path", "Workspace Mode", "Summary");
    String separator =
        "    "
            + "-".repeat(labelIdWidth)
            + "  "
            + "-".repeat(labelPathWidth)
            + "  "
            + "-".repeat(labelWorkspaceModeWidth)
            + "  "
            + "-".repeat("Summary".length());
    String body =
        shippedExamples.stream()
            .map(
                example ->
                    formatExampleLine(
                        example, labelIdWidth, labelPathWidth, labelWorkspaceModeWidth))
            .map(line -> line + "\n")
            .collect(java.util.stream.Collectors.joining())
            .stripTrailing();
    return header + "\n" + separator + "\n" + body;
  }

  private static String formatExampleLine(
      ShippedExampleEntry example, int idWidth, int pathWidth, int workspaceModeWidth) {
    String line =
        ("    %-" + idWidth + "s  %-" + pathWidth + "s  %-" + workspaceModeWidth + "s  %s")
            .formatted(
                example.id(),
                example.suggestedRequestPath(),
                example.workspaceMode().name(),
                example.summary());
    if (example.requiredPaths().isEmpty()) {
      return line;
    }
    return line + "\n" + "      requiredPaths: " + String.join(", ", example.requiredPaths());
  }

  static String formatTaskStarters(List<TaskEntry> tasks) {
    int idWidth = tasks.stream().mapToInt(task -> task.id().length()).max().orElse(0);
    int pathWidth =
        tasks.stream()
            .mapToInt(task -> task.starter().suggestedRequestPath().length())
            .max()
            .orElse(0);
    int workspaceModeWidth =
        tasks.stream()
            .mapToInt(task -> task.starter().workspaceMode().name().length())
            .max()
            .orElse(0);
    int labelIdWidth = Math.max(idWidth, "Task Id".length());
    int labelPathWidth = Math.max(pathWidth, "Starter Request Path".length());
    int labelWorkspaceModeWidth = Math.max(workspaceModeWidth, "Workspace Mode".length());
    String header =
        ("    %-"
                + labelIdWidth
                + "s  %-"
                + labelPathWidth
                + "s  %-"
                + labelWorkspaceModeWidth
                + "s  %s")
            .formatted("Task Id", "Starter Request Path", "Workspace Mode", "Summary");
    String separator =
        "    "
            + "-".repeat(labelIdWidth)
            + "  "
            + "-".repeat(labelPathWidth)
            + "  "
            + "-".repeat(labelWorkspaceModeWidth)
            + "  "
            + "-".repeat("Summary".length());
    List<String> rows = new ArrayList<>();
    for (TaskEntry task : tasks) {
      String line =
          ("    %-"
                  + labelIdWidth
                  + "s  %-"
                  + labelPathWidth
                  + "s  %-"
                  + labelWorkspaceModeWidth
                  + "s  %s")
              .formatted(
                  task.id(),
                  task.starter().suggestedRequestPath(),
                  task.starter().workspaceMode().name(),
                  task.narrative().summary());
      rows.add(line);
      if (!task.starter().requiredPaths().isEmpty()) {
        rows.add("      requiredPaths: " + String.join(", ", task.starter().requiredPaths()));
      }
    }
    return header + "\n" + separator + "\n" + String.join("\n", rows);
  }

  static String renderCoordinateSystems(CliSurface.CliTableSection section) {
    int leftWidth =
        Math.max(
            section.leftHeader().length(),
            section.entries().stream().mapToInt(entry -> entry.pattern().length()).max().orElse(0));
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
    return section.label() + ":\n" + indentLinesWrapped(section.lines(), 2);
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

  static String renderDiscovery(
      CliSurface.CliDiscoverySection section, String discoveryExamples, String taskStarters) {
    return section.label()
        + ":\n"
        + "  Discovery commands:\n"
        + indentLinesWrapped(section.lines(), 4)
        + "\n"
        + "  "
        + section.builtInExamplesLabel()
        + " catalog:\n"
        + discoveryExamples
        + "\n"
        + "  "
        + section.printOneExampleLabel()
        + ":\n"
        + "    "
        + section.printOneExampleCommand()
        + "\n"
        + "  Task starter catalog:\n"
        + taskStarters
        + "\n"
        + "  Print one task starter:\n"
        + "    gridgrind --print-task-plan --lookup DASHBOARD --response task-plan.json\n"
        + renderWorkflows(
            new CliSurface.CliWorkflowSection("Discovery guidance", section.guidanceEntries()));
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
                        + wrappedText(entry.description(), "    ", "    ", HELP_TEXT_WIDTH))
            .map(block -> block + "\n")
            .collect(java.util.stream.Collectors.joining())
            .stripTrailing();
  }

  private static String indentLinesWrapped(List<String> lines, int indentSpaces) {
    String indent = " ".repeat(indentSpaces);
    return lines.stream()
        .map(line -> wrappedIndentedLine(line, indent))
        .map(block -> block + "\n")
        .collect(java.util.stream.Collectors.joining())
        .stripTrailing();
  }

  private static String wrappedIndentedLine(String line, String indent) {
    java.util.regex.Matcher orderedMatcher =
        Pattern.compile("^(\\d+\\.)\\s+(.*)$").matcher(line.trim());
    if (orderedMatcher.matches()) {
      String firstPrefix = indent + orderedMatcher.group(1) + " ";
      String continuationPrefix = " ".repeat(firstPrefix.length());
      return wrappedText(orderedMatcher.group(2), firstPrefix, continuationPrefix, HELP_TEXT_WIDTH);
    }
    java.util.regex.Matcher bulletMatcher =
        Pattern.compile("^([-*])\\s+(.*)$").matcher(line.trim());
    if (bulletMatcher.matches()) {
      String firstPrefix = indent + bulletMatcher.group(1) + " ";
      String continuationPrefix = " ".repeat(firstPrefix.length());
      return wrappedText(bulletMatcher.group(2), firstPrefix, continuationPrefix, HELP_TEXT_WIDTH);
    }
    return wrappedText(line, indent, indent, HELP_TEXT_WIDTH);
  }

  private static String wrappedDefinition(CliSurface.DefinitionEntry entry, int width) {
    String prefix = ("  %-" + width + "s  ").formatted(entry.label() + ":");
    return wrappedText(entry.value(), prefix, " ".repeat(prefix.length()), HELP_TEXT_WIDTH);
  }

  private static String wrappedText(
      String text, String firstPrefix, String continuationPrefix, int width) {
    Objects.requireNonNull(text, "text must not be null");
    Objects.requireNonNull(firstPrefix, "firstPrefix must not be null");
    Objects.requireNonNull(continuationPrefix, "continuationPrefix must not be null");
    String normalizedText = text.trim();
    StringBuilder builder = new StringBuilder(firstPrefix);
    String currentPrefix = firstPrefix;
    int lineLength = currentPrefix.length();
    java.util.regex.Matcher matcher = WHITESPACE_PATTERN.matcher(normalizedText);
    int cursor = 0;
    while (cursor < normalizedText.length()) {
      int nextBoundary = normalizedText.length();
      if (matcher.find(cursor)) {
        nextBoundary = matcher.start();
      }
      String word = normalizedText.substring(cursor, nextBoundary);
      int separatorWidth = lineLength == currentPrefix.length() ? 0 : 1;
      if (lineLength + separatorWidth + word.length() > width
          && lineLength > currentPrefix.length()) {
        builder.append('\n').append(continuationPrefix).append(word);
        currentPrefix = continuationPrefix;
        lineLength = continuationPrefix.length() + word.length();
      } else {
        if (separatorWidth == 1) {
          builder.append(' ');
          lineLength++;
        }
        builder.append(word);
        lineLength += word.length();
      }
      cursor = nextBoundary;
      while (cursor < normalizedText.length()
          && Character.isWhitespace(normalizedText.charAt(cursor))) {
        cursor++;
      }
    }
    return builder.toString();
  }

  private static String replacePlaceholders(String value, String containerTag) {
    return value.replace("{{CONTAINER_TAG}}", containerTag);
  }
}
