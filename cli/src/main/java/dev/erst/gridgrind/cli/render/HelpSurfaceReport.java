package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** CLI-owned structured help metadata rendered by the command-line transport. */
public record HelpSurfaceReport(
    CliSection usage,
    CliWorkflowSection workflows,
    CliSection execution,
    CliDefinitionSection limits,
    CliSection request,
    CliDefinitionSection fileWorkflow,
    CliTableSection coordinateSystems,
    CliTemplateSection minimalValidRequest,
    CliCommandExample stdinExample,
    CliCommandExample dockerFileExample,
    CliDiscoverySection discovery,
    CliReferenceSection docs,
    CliDefinitionSection flags,
    String standardInputRequiresRequestMessage) {
  public HelpSurfaceReport {
    Objects.requireNonNull(usage, "usage must not be null");
    Objects.requireNonNull(workflows, "workflows must not be null");
    Objects.requireNonNull(execution, "execution must not be null");
    Objects.requireNonNull(limits, "limits must not be null");
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(fileWorkflow, "fileWorkflow must not be null");
    Objects.requireNonNull(coordinateSystems, "coordinateSystems must not be null");
    Objects.requireNonNull(minimalValidRequest, "minimalValidRequest must not be null");
    Objects.requireNonNull(stdinExample, "stdinExample must not be null");
    Objects.requireNonNull(dockerFileExample, "dockerFileExample must not be null");
    Objects.requireNonNull(discovery, "discovery must not be null");
    Objects.requireNonNull(docs, "docs must not be null");
    Objects.requireNonNull(flags, "flags must not be null");
    standardInputRequiresRequestMessage =
        TextValidation.requireNonBlank(
            standardInputRequiresRequestMessage, "standardInputRequiresRequestMessage");
  }

  /** One labeled bullet-list or command-list section. */
  public record CliSection(String label, List<String> lines) {
    public CliSection {
      label = TextValidation.requireNonBlank(label, "label");
      lines = HelpValidation.copyStrings(lines, "lines");
    }
  }

  /** One labeled key/value section whose renderer controls alignment. */
  public record CliDefinitionSection(String label, List<DefinitionEntry> entries) {
    public CliDefinitionSection {
      label = TextValidation.requireNonBlank(label, "label");
      Objects.requireNonNull(entries, "entries must not be null");
      entries =
          entries.stream()
              .map(entry -> Objects.requireNonNull(entry, "entries must not contain nulls"))
              .toList();
    }
  }

  /** One labeled key/value pair inside a definition section. */
  public record DefinitionEntry(String label, String value) {
    public DefinitionEntry {
      label = TextValidation.requireNonBlank(label, "label");
      value = TextValidation.requireNonBlank(value, "value");
    }
  }

  /** One labeled fixed-width table section. */
  public record CliTableSection(
      String label, String leftHeader, String rightHeader, List<CoordinateSystemEntry> entries) {
    public CliTableSection {
      label = TextValidation.requireNonBlank(label, "label");
      leftHeader = TextValidation.requireNonBlank(leftHeader, "leftHeader");
      rightHeader = TextValidation.requireNonBlank(rightHeader, "rightHeader");
      Objects.requireNonNull(entries, "entries must not be null");
      entries =
          entries.stream()
              .map(entry -> Objects.requireNonNull(entry, "entries must not contain nulls"))
              .toList();
    }
  }

  /** One coordinate-system row published in CLI help. */
  public record CoordinateSystemEntry(String pattern, String convention) {
    public CoordinateSystemEntry {
      pattern = TextValidation.requireNonBlank(pattern, "pattern");
      convention = TextValidation.requireNonBlank(convention, "convention");
    }
  }

  /** One labeled workflow section for first-contact operator guidance. */
  public record CliWorkflowSection(String label, List<WorkflowEntry> entries) {
    public CliWorkflowSection {
      label = TextValidation.requireNonBlank(label, "label");
      Objects.requireNonNull(entries, "entries must not be null");
      entries =
          entries.stream()
              .map(entry -> Objects.requireNonNull(entry, "entries must not contain nulls"))
              .toList();
    }
  }

  /** One titled workflow entry made of ordered guidance lines. */
  public record WorkflowEntry(String title, List<String> lines) {
    public WorkflowEntry {
      title = TextValidation.requireNonBlank(title, "title");
      lines = HelpValidation.copyStrings(lines, "lines");
    }
  }

  /** One label for the generated minimal valid request block. */
  public record CliTemplateSection(String label) {
    public CliTemplateSection {
      label = TextValidation.requireNonBlank(label, "label");
    }
  }

  /** One labeled CLI command example with optional explanatory prose and placeholder expansion. */
  public record CliCommandExample(
      String label, List<String> commandLines, Optional<String> description) {
    public CliCommandExample {
      label = TextValidation.requireNonBlank(label, "label");
      commandLines = HelpValidation.copyStrings(commandLines, "commandLines");
      Objects.requireNonNull(description, "description must not be null");
      if (description.isPresent() && description.orElseThrow().isBlank()) {
        throw new IllegalArgumentException("description must not be blank");
      }
    }
  }

  /** Discovery metadata plus the canonical featured built-in example id. */
  public record CliDiscoverySection(
      String label,
      List<String> lines,
      String builtInExamplesLabel,
      String printOneExampleLabel,
      List<WorkflowEntry> guidanceEntries,
      String printOneExampleCommand) {
    public CliDiscoverySection {
      label = TextValidation.requireNonBlank(label, "label");
      lines = HelpValidation.copyStrings(lines, "lines");
      builtInExamplesLabel =
          TextValidation.requireNonBlank(builtInExamplesLabel, "builtInExamplesLabel");
      printOneExampleLabel =
          TextValidation.requireNonBlank(printOneExampleLabel, "printOneExampleLabel");
      Objects.requireNonNull(guidanceEntries, "guidanceEntries must not be null");
      guidanceEntries =
          guidanceEntries.stream()
              .map(entry -> Objects.requireNonNull(entry, "guidanceEntries must not contain nulls"))
              .toList();
      printOneExampleCommand =
          TextValidation.requireNonBlank(printOneExampleCommand, "printOneExampleCommand");
    }
  }

  /** One label plus one ordered set of document references rooted at the docs path. */
  public record CliReferenceSection(String label, List<ReferenceEntry> entries) {
    public CliReferenceSection {
      label = TextValidation.requireNonBlank(label, "label");
      Objects.requireNonNull(entries, "entries must not be null");
      entries =
          entries.stream()
              .map(entry -> Objects.requireNonNull(entry, "entries must not contain nulls"))
              .toList();
    }
  }

  /** One labeled document reference with its docs-relative markdown path and one purpose line. */
  public record ReferenceEntry(String label, String relativePath, String description) {
    public ReferenceEntry {
      label = TextValidation.requireNonBlank(label, "label");
      relativePath = TextValidation.requireNonBlank(relativePath, "relativePath");
      description = TextValidation.requireNonBlank(description, "description");
    }
  }
}
