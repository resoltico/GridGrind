package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.catalog.GridGrindContainerRuntimeText;
import dev.erst.gridgrind.contract.catalog.GridGrindInspectionContractText;
import java.util.List;
import java.util.Optional;

/** Discovery, examples, and reference-oriented CLI help sections. */
final class GridGrindCliSurfaceGuidanceSections {
  private GridGrindCliSurfaceGuidanceSections() {}

  static CliSurface.CliTableSection coordinateSystems() {
    return new CliSurface.CliTableSection(
        "Coordinate Systems",
        "Pattern",
        "Convention / Example",
        List.of(
            new CliSurface.CoordinateSystemEntry("address", "A1 cell address, e.g. B3"),
            new CliSurface.CoordinateSystemEntry("range", "A1 rectangular range, e.g. A1:C4"),
            new CliSurface.CoordinateSystemEntry("*RowIndex", "zero-based, e.g. 0 = Excel row 1"),
            new CliSurface.CoordinateSystemEntry(
                "*ColumnIndex", "zero-based, e.g. 0 = Excel column A"),
            new CliSurface.CoordinateSystemEntry(
                "first/last pairs", "inclusive zero-based bands.")));
  }

  static CliSurface.CliTemplateSection minimalValidRequest() {
    return new CliSurface.CliTemplateSection("Minimal Valid Request");
  }

  static CliSurface.CliCommandExample stdinExample() {
    return new CliSurface.CliCommandExample(
        "Stdin Example",
        List.of("gridgrind --print-request-template | gridgrind --execution-root ."),
        Optional.empty());
  }

  static CliSurface.CliCommandExample dockerExample() {
    return new CliSurface.CliCommandExample(
        "Docker Example",
        List.of(
            "docker run --rm -i \\",
            "  " + GridGrindContainerRuntimeText.dockerMountedWorkdirUserArgument() + " \\",
            "  " + GridGrindContainerRuntimeText.dockerMountedWorkdirVolumeArgument() + " \\",
            "  {{CONTAINER_TAG}} \\",
            "  --request request.json \\",
            "  --response response.json"),
        Optional.of(
            GridGrindContainerRuntimeText.dockerMountedWorkdirSummary()
                + " From a repository checkout, build the same runtime surface with"
                + " 'docker buildx build --load -t"
                + " gridgrind-local .' and replace {{CONTAINER_TAG}} with"
                + " 'gridgrind-local'."));
  }

  static CliSurface.CliDiscoverySection discovery() {
    return new CliSurface.CliDiscoverySection(
        "Discovery",
        List.of(
            "gridgrind --print-request-template --response request.json",
            "gridgrind --doctor-request --request request.json --response doctor.json",
            "gridgrind --print-recipe-catalog --response recipes.json",
            "gridgrind --print-recipe --lookup <id> --response recipe.json",
            "gridgrind --print-recipe-keyword-match --query \"monthly sales dashboard with charts\""
                + " --response recipe-keyword-match.json",
            "gridgrind --print-protocol-catalog --response protocol-index.json"),
        "Built-in generated examples",
        "Print one recipe",
        List.of(
            new CliSurface.WorkflowEntry(
                "Example portability",
                List.of(
                    "SELF_CONTAINED starters execute from a blank working directory.",
                    "REQUIRES_EXAMPLE_ASSETS starters require copied asset paths beside the"
                        + " request file; requiredWorkspacePaths names those paths directly.",
                    GridGrindInspectionContractText.workbookFindingsDiscoverySummary()
                        + " Include it in any diagnostic plan with persistence.type=NONE.")),
            new CliSurface.WorkflowEntry(
                "Task starters",
                List.of(
                    "The unified recipe catalog includes high-level office-work task"
                        + " starters composed from exact protocol capabilities.",
                    "Each task-starter entry publishes requestFileName, workspaceMode, and"
                        + " requiredWorkspacePaths so agents can decide whether one"
                        + " recipe is self-contained before printing it with"
                        + " --print-recipe.",
                    "Use --print-recipe-catalog --lookup <id> when you need the richer"
                        + " view-specific detail payload, including the exact runnable"
                        + " request profile for that recipe.")),
            new CliSurface.WorkflowEntry(
                "Protocol catalog search",
                List.of(
                    "The protocol catalog remains the authoritative execution contract: it"
                        + " lists each field, whether it is required, and the nested/plain"
                        + " type group accepted by polymorphic fields such as target,"
                        + " action, query, value, style, and scope.",
                    "The bare --print-protocol-catalog output is intentionally compact:"
                        + " it lists requestTypeId, group ids, and lookup namespace forms."
                        + " Rerun --lookup for one scoped top-level category, nested type"
                        + " group, plain type group, or exact entry payload."
                        + " Shared reusable notes such as request-owned path resolution"
                        + " live on those scoped lookup payloads, not on the bare index.",
                    "Search output is summary-first: it lists ids, summaries, related entry"
                        + " ids, and supporting ids. Rerun --lookup when you need one full"
                        + " entry or type-group definition.",
                    "Lookup forms stay stable:"
                        + " bare <id> resolves one globally unique top-level type,"
                        + " bare top-level group names such as mutationActionTypes"
                        + " resolve one category array,"
                        + " bare support-group names such as cellInputTypes or"
                        + " executionPolicyInputType resolve one nested/plain group,"
                        + " nestedTypes:<group> and plainTypes:<group> resolve one"
                        + " explicit support-group namespace,"
                        + " and <topLevelGroup>:<id> resolves one qualified top-level"
                        + " type."))),
        "gridgrind --print-recipe --lookup "
            + GridGrindShippedExamples.find("BUDGET").orElseThrow().id()
            + " --response example.json");
  }

  static CliSurface.CliReferenceSection docs() {
    return new CliSurface.CliReferenceSection(
        "Docs",
        List.of(
            new CliSurface.ReferenceEntry(
                "Quick reference",
                "docs/QUICK_REFERENCE.md",
                "Synopsis-first command cheat sheet for first-contact lookup."),
            new CliSurface.ReferenceEntry(
                "Operations reference",
                "docs/OPERATIONS.md",
                "Longer operator workflows, artifact handling, and runtime/container use."),
            new CliSurface.ReferenceEntry(
                "Error reference",
                "docs/ERRORS.md",
                "Problem codes, failure interpretation, and recovery guidance.")));
  }
}
