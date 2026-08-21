package dev.erst.gridgrind.cli.discovery;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import java.util.LinkedHashSet;
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

    assertEquals("requestFileName must end with .json", failure.getMessage());
  }

  @Test
  void shippedExampleRejectsBlankDiscoveryFields() {
    IllegalArgumentException blankIdFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new GridGrindShippedExamples.ShippedExample(
                    " ",
                    "workbook-health-request.json",
                    "summary",
                    GridGrindProtocolCatalog.requestTemplate()));
    assertEquals("id must not be blank", blankIdFailure.getMessage());

    IllegalArgumentException blankSummaryFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new GridGrindShippedExamples.ShippedExample(
                    "WORKBOOK_HEALTH",
                    "workbook-health-request.json",
                    " ",
                    GridGrindProtocolCatalog.requestTemplate()));
    assertEquals("summary must not be blank", blankSummaryFailure.getMessage());
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
                    RecipeAdvisory.SELF_CONTAINED,
                    java.util.List.of()));

    assertEquals("requestFileName must end with .json", failure.getMessage());
  }

  @Test
  void shippedExampleEntryRejectsRepositoryPathsUsingWindowsSeparators() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new ShippedExampleEntry(
                    "WORKBOOK_HEALTH",
                    "examples\\workbook-health-request.json",
                    "summary",
                    RecipeAdvisory.SELF_CONTAINED,
                    java.util.List.of()));

    assertEquals(
        "requestFileName must be one portable file name, not a repository path",
        failure.getMessage());
  }

  @Test
  void shippedExampleEntryCarriesOnlyPublicSummaryFields() {
    ShippedExampleEntry entry =
        new ShippedExampleEntry(
            "WORKBOOK_HEALTH",
            "workbook-health-request.json",
            "summary",
            RecipeAdvisory.SELF_CONTAINED,
            java.util.List.of());
    assertEquals("WORKBOOK_HEALTH", entry.id());
    assertEquals("workbook-health-request.json", entry.requestFileName());
    assertEquals("summary", entry.summary());
    assertEquals(RecipeAdvisory.SELF_CONTAINED, entry.advisory());
    assertEquals(java.util.List.of(), entry.requiredWorkspacePaths());
  }

  @Test
  void shippedExampleEntryRejectsMissingRequiredPaths() {
    NullPointerException failure =
        assertThrows(
            NullPointerException.class,
            () ->
                new ShippedExampleEntry(
                    "WORKBOOK_HEALTH",
                    "workbook-health-request.json",
                    "summary",
                    RecipeAdvisory.SELF_CONTAINED,
                    null));

    assertEquals("requiredWorkspacePaths must not be null", failure.getMessage());
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
        GridGrindShippedExamples.repositoryAssetBackedExamples().stream()
            .map(GridGrindShippedExamples::requirementsFor)
            .allMatch(entry -> !entry.requiredWorkspacePaths().isEmpty()));
  }

  @Test
  void internalCatalogEntryAndExampleRequirementsGuardsStayDefensive() {
    IllegalStateException missingRequirementsFailure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindShippedExamples.requirementsFor(
                    new GridGrindShippedExamples.ShippedExample(
                        "NO_REQUIREMENTS",
                        "no-requirements.json",
                        "summary",
                        GridGrindProtocolCatalog.requestTemplate())));
    assertEquals(
        "Missing shipped-example requirements for NO_REQUIREMENTS",
        missingRequirementsFailure.getMessage());

    NullPointerException missingRequiredPathsFailure =
        assertThrows(
            NullPointerException.class,
            () ->
                new GridGrindShippedExamples.ExampleRequirements(
                    RecipeAdvisory.SELF_CONTAINED, null));
    assertEquals(
        "requiredWorkspacePaths must not be null", missingRequiredPathsFailure.getMessage());
  }

  @Test
  void exampleLookupAndAdvisoryAccessorsStayBoundToExampleRecipesOnly() {
    NullPointerException nullLookup =
        assertThrows(NullPointerException.class, () -> GridGrindShippedExamples.find(null));
    assertEquals("id must not be null", nullLookup.getMessage());
    assertTrue(GridGrindShippedExamples.find("WORKBOOK_HEALTH").isPresent());
    assertTrue(GridGrindShippedExamples.find("DASHBOARD").isEmpty());

    NullPointerException nullId =
        assertThrows(NullPointerException.class, () -> GridGrindShippedExamples.advisoryFor(null));
    assertEquals("id must not be null", nullId.getMessage());

    assertEquals(
        java.util.Optional.of(RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS),
        GridGrindShippedExamples.advisoryFor("PACKAGE_SECURITY_INSPECTION"));
    assertEquals(java.util.Optional.empty(), GridGrindShippedExamples.advisoryFor("DASHBOARD"));
    assertEquals(
        java.util.Optional.empty(), GridGrindShippedExamples.advisoryFor("NO_SUCH_EXAMPLE"));
  }

  @Test
  void requirementsLookupRejectsTaskRecipeIdsMasqueradingAsExamples() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () ->
                GridGrindShippedExamples.requirementsFor(
                    new GridGrindShippedExamples.ShippedExample(
                        "DASHBOARD",
                        "dashboard-request.json",
                        "summary",
                        GridGrindProtocolCatalog.requestTemplate())));

    assertEquals("Missing shipped-example requirements for DASHBOARD", failure.getMessage());
  }
}
