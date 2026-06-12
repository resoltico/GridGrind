package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.ExampleWorkspaceMode;
import dev.erst.gridgrind.cli.discovery.GridGrindTaskCatalog;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogGroupIndex;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogIndexReport;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogLookupNamespace;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogSearchHit;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogSearchReport;
import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.catalog.CatalogSearchMatch;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

/** Shared discovery, warning, and search-summary helpers for CLI catalog commands. */
final class CliCatalogCommandSupport {
  private CliCatalogCommandSupport() {}

  static String unknownExampleMessage(String exampleId) {
    return suggestedExampleId(exampleId)
        .map(
            suggestion ->
                "Unknown example: "
                    + exampleId
                    + ". Example ids use stable upper-case tokens; did you mean "
                    + suggestion
                    + "? Run gridgrind --print-example-catalog to list valid ids.")
        .orElse(
            "Unknown example: "
                + exampleId
                + ". Run gridgrind --print-example-catalog to list valid ids.");
  }

  static void emitExamplePortabilityWarning(
      GridGrindShippedExamples.ShippedExample example, OutputStream stderr) throws IOException {
    Objects.requireNonNull(example, "example must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    if (GridGrindShippedExamples.workspaceModeFor(example.id()).orElseThrow()
        != ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS) {
      return;
    }
    var requirements = GridGrindShippedExamples.requirementsFor(example);
    String requiredPaths = String.join(", ", requirements.requiredPaths());
    stderr.write(
        ("Printed example "
                + example.id()
                + " requires copied asset paths beside the request file before execution;"
                + " required paths: "
                + requiredPaths
                + ". Inspect --print-example-catalog or --help-guidance for portability details.\n")
            .getBytes(StandardCharsets.UTF_8));
    stderr.flush();
  }

  static void emitTaskStarterPortabilityWarning(
      dev.erst.gridgrind.cli.discovery.TaskEntry task, OutputStream stderr) throws IOException {
    Objects.requireNonNull(task, "task must not be null");
    Objects.requireNonNull(stderr, "stderr must not be null");
    if (task.starter().workspaceMode() != ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS) {
      return;
    }
    String requiredPaths = String.join(", ", task.starter().requiredPaths());
    stderr.write(
        ("Printed task starter "
                + task.id()
                + " requires copied asset paths beside the request file before execution;"
                + " required paths: "
                + requiredPaths
                + ". Inspect --print-task-catalog or --help-guidance for starter portability"
                + " details.\n")
            .getBytes(StandardCharsets.UTF_8));
    stderr.flush();
  }

  static String unknownTaskMessage(String taskId) {
    return suggestedTaskId(taskId)
        .map(
            suggestion ->
                "Unknown task: "
                    + taskId
                    + ". Task ids use stable upper-case tokens; did you mean "
                    + suggestion
                    + "? Run gridgrind --print-task-catalog to list valid ids or"
                    + " gridgrind --print-task-keyword-match --query \"monthly sales dashboard\""
                    + " to discover a close task id before printing its task plan.")
        .orElse(
            "Unknown task: "
                + taskId
                + ". Run gridgrind --print-task-catalog to list valid ids or"
                + " gridgrind --print-task-keyword-match --query \"monthly sales dashboard\""
                + " to discover a close task id before printing its task plan.");
  }

  static String unknownOperationMessage(String operationId) {
    return "Unknown lookup id: "
        + operationId
        + ". Run gridgrind --print-protocol-catalog --search \"sheet layout\" or"
        + " gridgrind --print-protocol-catalog to discover valid lookup ids.";
  }

  static ProtocolCatalogIndexReport protocolCatalogIndexReport() {
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    return new ProtocolCatalogIndexReport(
        catalog.protocolVersion(),
        catalog.discriminatorField(),
        catalog.requestType().id(),
        catalog.topLevelGroups().stream()
            .map(group -> new ProtocolCatalogGroupIndex(group.group(), typeIds(group.types())))
            .toList(),
        catalog.nestedTypes().stream()
            .map(group -> new ProtocolCatalogGroupIndex(group.group(), typeIds(group.types())))
            .toList(),
        catalog.plainTypes().stream()
            .map(group -> new ProtocolCatalogGroupIndex(group.group(), List.of(group.type().id())))
            .toList(),
        List.of(
            new ProtocolCatalogLookupNamespace(
                "<topLevelGroup>:<id>",
                "Resolve one top-level protocol type such as mutationActionTypes:SET_CELL."),
            new ProtocolCatalogLookupNamespace(
                "nestedTypes:<group>",
                "Resolve one nested tagged-union group such as nestedTypes:cellInputTypes."),
            new ProtocolCatalogLookupNamespace(
                "plainTypes:<group>",
                "Resolve one plain record group such as plainTypes:sheetSummaryReport."),
            new ProtocolCatalogLookupNamespace(
                "<id>",
                "Resolve one unqualified top-level type id only when it is globally unique.")));
  }

  static ProtocolCatalogSearchReport summarizedSearchReport(String query) {
    var result = GridGrindProtocolCatalog.searchCatalog(query);
    return new ProtocolCatalogSearchReport(
        result.protocolVersion(),
        result.query(),
        result.matches().stream().map(CliCatalogCommandSupport::summarizedSearchHit).toList());
  }

  private static ProtocolCatalogSearchHit summarizedSearchHit(CatalogSearchMatch match) {
    return new ProtocolCatalogSearchHit(
        match.catalogGroup(),
        match.lookupId(),
        match.qualifiedId(),
        match.kind(),
        match.summary(),
        match.relatedEntryIds(),
        match.supportingMatches().stream().map(CatalogSearchMatch::qualifiedId).toList());
  }

  private static List<String> typeIds(List<dev.erst.gridgrind.contract.catalog.TypeEntry> types) {
    return types.stream().map(dev.erst.gridgrind.contract.catalog.TypeEntry::id).toList();
  }

  private static Optional<String> suggestedTaskId(String taskId) {
    String normalizedTaskId = normalizeLookupToken(taskId);
    return GridGrindTaskCatalog.catalog().tasks().stream()
        .map(dev.erst.gridgrind.cli.discovery.TaskEntry::id)
        .filter(
            candidate ->
                candidate.equalsIgnoreCase(taskId)
                    || normalizeLookupToken(candidate).equals(normalizedTaskId))
        .findFirst();
  }

  private static Optional<String> suggestedExampleId(String exampleId) {
    String normalizedExampleId = normalizeLookupToken(exampleId);
    return GridGrindShippedExamples.examples().stream()
        .filter(
            example ->
                example.id().equalsIgnoreCase(exampleId)
                    || example.requestFileName().equalsIgnoreCase(exampleId)
                    || exampleStem(example.requestFileName()).equalsIgnoreCase(exampleId)
                    || normalizeLookupToken(exampleStem(example.requestFileName()))
                        .equals(normalizedExampleId))
        .map(GridGrindShippedExamples.ShippedExample::id)
        .findFirst();
  }

  private static String normalizeLookupToken(String value) {
    return value.toUpperCase(Locale.ROOT).replace('-', '_').replace(' ', '_');
  }

  private static String exampleStem(String fileName) {
    return fileName.substring(0, fileName.length() - 5);
  }
}
