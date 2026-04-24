package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;
import org.junit.jupiter.api.Test;

/** Tests for generated-example metadata validation and accessors. */
class GridGrindShippedExamplesTest {
  @Test
  void shippedExampleRejectsLowerCaseDiscoveryIds() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new GridGrindShippedExamples.ShippedExample(
                    "workbook_health",
                    "workbook-health-request.json",
                    "summary",
                    GridGrindProtocolCatalog.requestTemplate()));

    assertEquals("id must use upper-case discovery tokens", failure.getMessage());
  }

  @Test
  void shippedExampleRejectsNonJsonFileNames() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new GridGrindShippedExamples.ShippedExample(
                    "WORKBOOK_HEALTH",
                    "workbook-health-request.txt",
                    "summary",
                    GridGrindProtocolCatalog.requestTemplate()));

    assertEquals("fileName must end with .json", failure.getMessage());
  }

  @Test
  void shippedExampleEntryRejectsNonJsonFileNames() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new ShippedExampleEntry(
                    "WORKBOOK_HEALTH",
                    "workbook-health-request.txt",
                    "summary",
                    ShippedExampleEntry.WorkspaceMode.BLANK_WORKSPACE,
                    List.of()));

    assertEquals("fileName must end with .json", failure.getMessage());
  }

  @Test
  void shippedExampleEntryValidatesWorkspaceModeAndRequiredPathsTogether() {
    IllegalArgumentException blankWorkspaceFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new ShippedExampleEntry(
                    "WORKBOOK_HEALTH",
                    "workbook-health-request.json",
                    "summary",
                    ShippedExampleEntry.WorkspaceMode.BLANK_WORKSPACE,
                    List.of("examples/source-backed-input-assets/title.txt")));
    assertEquals(
        "requiredPaths must be empty when workspaceMode is BLANK_WORKSPACE",
        blankWorkspaceFailure.getMessage());

    IllegalArgumentException repositoryAssetsFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new ShippedExampleEntry(
                    "CUSTOM_XML",
                    "custom-xml-request.json",
                    "summary",
                    ShippedExampleEntry.WorkspaceMode.REPOSITORY_ASSETS,
                    List.of()));
    assertEquals(
        "requiredPaths must not be empty when workspaceMode is REPOSITORY_ASSETS",
        repositoryAssetsFailure.getMessage());
  }

  @Test
  void shippedExampleEntryNormalizesNullRequiredPathsAndRejectsTrailingSlashes() {
    ShippedExampleEntry blankWorkspaceEntry =
        new ShippedExampleEntry(
            "WORKBOOK_HEALTH",
            "workbook-health-request.json",
            "summary",
            ShippedExampleEntry.WorkspaceMode.BLANK_WORKSPACE,
            null);
    assertEquals(List.of(), blankWorkspaceEntry.requiredPaths());

    IllegalArgumentException trailingSlashFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new ShippedExampleEntry(
                    "CUSTOM_XML",
                    "custom-xml-request.json",
                    "summary",
                    ShippedExampleEntry.WorkspaceMode.REPOSITORY_ASSETS,
                    List.of("examples/custom-xml-assets/")));
    assertEquals("requiredPath must not end with /", trailingSlashFailure.getMessage());
  }

  @Test
  void selfContainedAndRepositoryAssetBackedPartitionsCoverEveryBuiltInExampleExactlyOnce() {
    Set<String> allIds =
        GridGrindShippedExamples.examples().stream()
            .map(GridGrindShippedExamples.ShippedExample::id)
            .collect(Collectors.toCollection(LinkedHashSet::new));
    Set<String> selfContainedIds =
        GridGrindShippedExamples.selfContainedExamples().stream()
            .map(GridGrindShippedExamples.ShippedExample::id)
            .collect(Collectors.toCollection(LinkedHashSet::new));
    Set<String> repositoryAssetIds =
        GridGrindShippedExamples.repositoryAssetBackedExamples().stream()
            .map(GridGrindShippedExamples.ShippedExample::id)
            .collect(Collectors.toCollection(LinkedHashSet::new));
    Set<String> partitionIds =
        Stream.concat(selfContainedIds.stream(), repositoryAssetIds.stream())
            .collect(Collectors.toCollection(LinkedHashSet::new));

    assertEquals(allIds, partitionIds);
    assertTrue(
        GridGrindShippedExamples.catalogEntries().stream()
            .filter(
                entry ->
                    entry.workspaceMode() == ShippedExampleEntry.WorkspaceMode.REPOSITORY_ASSETS)
            .allMatch(entry -> !entry.requiredPaths().isEmpty()));
  }

  @Test
  void internalCatalogEntryAndExampleRequirementsGuardsStayDefensive() {
    IllegalStateException missingRequirementsFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindShippedExamples.catalogEntry(
                    new GridGrindShippedExamples.ShippedExample(
                        "NO_REQUIREMENTS",
                        "no-requirements.json",
                        "summary",
                        GridGrindProtocolCatalog.requestTemplate())));
    assertEquals(
        "Missing shipped-example requirements for NO_REQUIREMENTS",
        missingRequirementsFailure.getMessage());

    GridGrindShippedExamples.ExampleRequirements requirements =
        new GridGrindShippedExamples.ExampleRequirements(
            ShippedExampleEntry.WorkspaceMode.BLANK_WORKSPACE, null);
    assertEquals(List.of(), requirements.requiredPaths());
  }
}
