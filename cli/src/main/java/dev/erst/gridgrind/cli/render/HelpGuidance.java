package dev.erst.gridgrind.cli.render;

import java.util.List;
import java.util.Optional;

/** Discovery, examples, and reference-oriented CLI help sections. */
final class HelpGuidance {
  private HelpGuidance() {}

  static HelpSurfaceReport.CliTableSection coordinateSystems() {
    return new HelpSurfaceReport.CliTableSection(
        "Coordinate Systems",
        "Pattern",
        "Convention / Example",
        List.of(
            new HelpSurfaceReport.CoordinateSystemEntry("address", "A1 cell address, e.g. B3"),
            new HelpSurfaceReport.CoordinateSystemEntry(
                "range", "A1 rectangular range, e.g. A1:C4"),
            new HelpSurfaceReport.CoordinateSystemEntry(
                "*RowIndex", "zero-based, e.g. 0 = Excel row 1"),
            new HelpSurfaceReport.CoordinateSystemEntry(
                "*ColumnIndex", "zero-based, e.g. 0 = Excel column A"),
            new HelpSurfaceReport.CoordinateSystemEntry(
                "first/last pairs", "inclusive zero-based bands.")));
  }

  static HelpSurfaceReport.CliTemplateSection minimalValidRequest() {
    return new HelpSurfaceReport.CliTemplateSection("Minimal Valid Request");
  }

  static HelpSurfaceReport.CliCommandExample stdinExample() {
    return new HelpSurfaceReport.CliCommandExample(
        "Standard input",
        List.of("gridgrind run --request - --execution-root . --grant grant.json < request.json"),
        java.util.Optional.of(
            "The request document and STANDARD_INPUT payloads cannot consume the same stream."));
  }

  static HelpSurfaceReport.CliCommandExample dockerExample() {
    return new HelpSurfaceReport.CliCommandExample(
        "Docker Example",
        List.of(
            "docker run --rm -i \\",
            "  " + ContainerHelp.dockerMountedWorkdirUserArgument() + " \\",
            "  " + ContainerHelp.dockerMountedWorkdirVolumeArgument() + " \\",
            "  {{CONTAINER_TAG}} \\",
            "  run --request request.json \\",
            "  --grant grant.json \\",
            "  --response response.json"),
        Optional.of(
            ContainerHelp.dockerMountedWorkdirSummary()
                + " From a repository checkout, build the same runtime surface with"
                + " 'docker buildx build --load -t"
                + " gridgrind-local .' and replace {{CONTAINER_TAG}} with"
                + " 'gridgrind-local'."));
  }

  static HelpSurfaceReport.CliDiscoverySection discovery() {
    return new HelpSurfaceReport.CliDiscoverySection(
        "Discovery",
        List.of(
            "gridgrind recipe list",
            "gridgrind recipe show BUDGET --response budget-request.json",
            "gridgrind recipe match \"monthly sales dashboard with charts\"",
            "gridgrind recipe materialize CUSTOM_XML --workspace custom-xml",
            "gridgrind catalog search \"chart title\"",
            "gridgrind schema request --response request.schema.json"),
        "Built-in recipes",
        "Show one recipe",
        List.of(
            new HelpSurfaceReport.WorkflowEntry(
                "Portability",
                List.of(
                    "SELF_CONTAINED recipes run from a blank workspace.",
                    "Recipes with requiredWorkspacePaths must be materialized into a new workspace.",
                    "Recipe and task metadata are also available through the core discovery API.")),
            new HelpSurfaceReport.WorkflowEntry(
                "Protocol",
                List.of(
                    "Catalog search returns summary-first matches; catalog show returns exact fields and constraints.",
                    "Use qualified group:id values when an ID appears in more than one group.",
                    "Use schema request, schema grant and schema response for the exchanged documents."))),
        "gridgrind recipe show BUDGET --response budget-request.json");
  }

  static HelpSurfaceReport.CliReferenceSection docs() {
    return new HelpSurfaceReport.CliReferenceSection(
        "Docs",
        List.of(
            new HelpSurfaceReport.ReferenceEntry(
                "Quick reference",
                "docs/QUICK_REFERENCE.md",
                "Synopsis-first command cheat sheet for first-contact lookup."),
            new HelpSurfaceReport.ReferenceEntry(
                "Operations reference",
                "docs/OPERATIONS.md",
                "Longer operator workflows, artifact handling, and runtime/container use."),
            new HelpSurfaceReport.ReferenceEntry(
                "Error reference",
                "docs/ERRORS.md",
                "Problem codes, failure interpretation, and recovery guidance.")));
  }
}
