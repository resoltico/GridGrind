package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Objects;

/** Contract-owned structured CLI help metadata rendered by thin downstream transports. */
public record CliSurface(
    CliSection usage,
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
  public CliSurface {
    Objects.requireNonNull(usage, "usage must not be null");
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
        CatalogRecordValidation.requireNonBlank(
            standardInputRequiresRequestMessage, "standardInputRequiresRequestMessage");
  }

  /** One labeled bullet-list or command-list section. */
  public record CliSection(String label, List<String> lines) {
    public CliSection {
      label = CatalogRecordValidation.requireNonBlank(label, "label");
      lines = CatalogRecordValidation.copyStrings(lines, "lines");
    }
  }

  /** One labeled key/value section whose renderer controls alignment. */
  public record CliDefinitionSection(String label, List<DefinitionEntry> entries) {
    public CliDefinitionSection {
      label = CatalogRecordValidation.requireNonBlank(label, "label");
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
      label = CatalogRecordValidation.requireNonBlank(label, "label");
      value = CatalogRecordValidation.requireNonBlank(value, "value");
    }
  }

  /** One labeled fixed-width table section. */
  public record CliTableSection(
      String label, String leftHeader, String rightHeader, List<CoordinateSystemEntry> entries) {
    public CliTableSection {
      label = CatalogRecordValidation.requireNonBlank(label, "label");
      leftHeader = CatalogRecordValidation.requireNonBlank(leftHeader, "leftHeader");
      rightHeader = CatalogRecordValidation.requireNonBlank(rightHeader, "rightHeader");
      Objects.requireNonNull(entries, "entries must not be null");
      entries =
          entries.stream()
              .map(entry -> Objects.requireNonNull(entry, "entries must not contain nulls"))
              .toList();
    }
  }

  /** One coordinate-system row published in CLI help and the machine-readable catalog. */
  public record CoordinateSystemEntry(String pattern, String convention) {
    public CoordinateSystemEntry {
      pattern = CatalogRecordValidation.requireNonBlank(pattern, "pattern");
      convention = CatalogRecordValidation.requireNonBlank(convention, "convention");
    }
  }

  /** One label for the generated minimal valid request block. */
  public record CliTemplateSection(String label) {
    public CliTemplateSection {
      label = CatalogRecordValidation.requireNonBlank(label, "label");
    }
  }

  /** One labeled CLI command example with optional explanatory prose and placeholder expansion. */
  public record CliCommandExample(String label, List<String> commandLines, String description) {
    public CliCommandExample {
      label = CatalogRecordValidation.requireNonBlank(label, "label");
      commandLines = CatalogRecordValidation.copyStrings(commandLines, "commandLines");
      if (description != null && description.isBlank()) {
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
      String protocolCatalogNote,
      String printOneExampleCommand) {
    public CliDiscoverySection {
      label = CatalogRecordValidation.requireNonBlank(label, "label");
      lines = CatalogRecordValidation.copyStrings(lines, "lines");
      builtInExamplesLabel =
          CatalogRecordValidation.requireNonBlank(builtInExamplesLabel, "builtInExamplesLabel");
      printOneExampleLabel =
          CatalogRecordValidation.requireNonBlank(printOneExampleLabel, "printOneExampleLabel");
      protocolCatalogNote =
          CatalogRecordValidation.requireNonBlank(protocolCatalogNote, "protocolCatalogNote");
      printOneExampleCommand =
          CatalogRecordValidation.requireNonBlank(printOneExampleCommand, "printOneExampleCommand");
    }
  }

  /** One label plus one ordered set of document references rooted at the docs path. */
  public record CliReferenceSection(String label, List<ReferenceEntry> entries) {
    public CliReferenceSection {
      label = CatalogRecordValidation.requireNonBlank(label, "label");
      Objects.requireNonNull(entries, "entries must not be null");
      entries =
          entries.stream()
              .map(entry -> Objects.requireNonNull(entry, "entries must not contain nulls"))
              .toList();
    }
  }

  /** One label plus one docs-relative markdown path. */
  public record ReferenceEntry(String label, String relativePath) {
    public ReferenceEntry {
      label = CatalogRecordValidation.requireNonBlank(label, "label");
      relativePath = CatalogRecordValidation.requireNonBlank(relativePath, "relativePath");
    }
  }
}
