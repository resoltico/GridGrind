package dev.erst.gridgrind.core.discovery;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.core.catalog.TypeEntry;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.HostArgument;
import dev.erst.gridgrind.core.result.CommandError;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct coverage for CLI discovery validation helpers and task-definition accessors. */
class CliDiscoveryValidationCoverageTest {
  @Test
  void validatesPrimitiveCopyHelpersAndNullGuards() {
    ProtocolVersion protocolVersion = ProtocolVersion.CURRENT;
    TypeEntry typeEntry = GridGrindProtocolCatalog.catalog().plainTypes().getFirst().type();
    ShippedExampleEntry exampleEntry = ShippedExamples.catalog().examples().getFirst();
    TaskEntry taskEntry = TaskCatalogs.catalog().tasks().getFirst();
    TaskPhase taskPhase = taskEntry.workflow().phases().getFirst();
    TaskCapabilityRef capabilityRef = taskPhase.capabilityRefs().getFirst();

    assertEquals(protocolVersion, DiscoveryValidation.requireProtocolVersion(protocolVersion));
    assertEquals(
        List.of("alpha", "beta"),
        DiscoveryValidation.copyStrings(List.of("alpha", "beta"), "values"));
    assertEquals(
        List.of(typeEntry), DiscoveryValidation.copyTypeEntries(List.of(typeEntry), "entries"));
    assertEquals(
        List.of(taskEntry), DiscoveryValidation.copyTaskEntries(List.of(taskEntry), "tasks"));
    assertEquals(
        List.of(taskPhase), DiscoveryValidation.copyTaskPhases(List.of(taskPhase), "phases"));
    assertEquals(
        List.of(capabilityRef),
        DiscoveryValidation.copyTaskCapabilityRefs(List.of(capabilityRef), "capabilityRefs"));
    assertEquals(
        List.of(exampleEntry),
        DiscoveryValidation.copyExampleEntries(List.of(exampleEntry), "examples"));
    assertEquals(
        Optional.of("value"),
        DiscoveryValidation.normalizeOptionalString(Optional.of("value"), "label"));
    assertEquals(
        Optional.empty(), DiscoveryValidation.normalizeOptionalString(Optional.empty(), "label"));
    assertEquals(
        Optional.of("value"),
        DiscoveryValidation.copyOptionalArbitraryString(Optional.of("value"), "query"));
    assertEquals(
        Optional.empty(),
        DiscoveryValidation.copyOptionalArbitraryString(Optional.empty(), "query"));

    assertEquals(
        "protocolVersion must not be null",
        assertThrows(
                NullPointerException.class, () -> DiscoveryValidation.requireProtocolVersion(null))
            .getMessage());
    assertEquals(
        "entries must not be null",
        assertThrows(
                NullPointerException.class,
                () -> DiscoveryValidation.copyTypeEntries(null, "entries"))
            .getMessage());
    assertEquals(
        "label must not be null",
        assertThrows(
                NullPointerException.class,
                () -> DiscoveryValidation.normalizeOptionalString(null, "label"))
            .getMessage());
    assertEquals(
        "label must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> DiscoveryValidation.normalizeOptionalString(Optional.of("   "), "label"))
            .getMessage());
    assertEquals(
        "query must not be null",
        assertThrows(
                NullPointerException.class,
                () -> DiscoveryValidation.copyOptionalArbitraryString(null, "query"))
            .getMessage());
    assertEquals(
        "entries must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    DiscoveryValidation.copyTypeEntries(
                        java.util.Arrays.asList((TypeEntry) null), "entries"))
            .getMessage());
  }

  @Test
  void taskCatalogDerivesCanonicalEntriesFromTheSharedRegistryAndRejectsInvalidLookupIds() {
    TaskCatalog catalog = TaskCatalogs.catalog();
    List<TaskEntry> entries = RecipeRegistry.taskEntries();

    assertFalse(entries.isEmpty());
    assertNotSame(entries, catalog.tasks());
    assertEquals(entries, catalog.tasks());
    assertTrue(TaskCatalogs.entryFor("DASHBOARD").isPresent());
    assertTrue(TaskCatalogs.entryFor("WORKBOOK_MAINTENANCE").isPresent());
    assertTrue(TaskCatalogs.entryFor("NO_SUCH_TASK").isEmpty());
    TaskCatalogs.validateCapabilityReferences(catalog);

    assertEquals(
        "id must not be null",
        assertThrows(NullPointerException.class, () -> TaskCatalogs.entryFor(null)).getMessage());
    assertEquals(
        "id must not be blank",
        assertThrows(IllegalArgumentException.class, () -> TaskCatalogs.entryFor("   "))
            .getMessage());
  }

  @Test
  void taskCatalogRejectsDanglingCapabilityReferences() {
    TaskEntry brokenTask =
        new TaskEntry(
            "BROKEN",
            List.of("office"),
            TaskTestFixtures.discoveryProfile("broken"),
            TaskTestFixtures.narrative("Broken task for direct validation coverage"),
            profile(),
            TaskTestFixtures.interactionProfile(),
            TaskStarterContract.selfContained("broken-request.json"),
            new TaskWorkflow(
                List.of(
                    new TaskPhase(
                        TaskPhasePurpose.AUTHOR,
                        "broken phase",
                        "exercise dangling capability validation",
                        List.of(new TaskCapabilityRef("mutationActionTypes", "NO_SUCH_ACTION")),
                        List.of())),
                List.of()));

    assertEquals(
        "Task BROKEN references unknown protocol capability mutationActionTypes:NO_SUCH_ACTION",
        assertThrows(
                IllegalStateException.class,
                () ->
                    TaskCatalogs.validateCapabilityReferences(
                        new TaskCatalog(ProtocolVersion.CURRENT, List.of(brokenTask))))
            .getMessage());
  }

  @Test
  void recipeCatalogValidationRejectsDuplicateIdsAndInvalidWorkspaceContracts() {
    assertTrue(RecipeCatalogs.entryFor("DASHBOARD").isPresent());
    assertTrue(RecipeCatalogs.lookupFor("DASHBOARD").isPresent());
    assertEquals(
        "id must not be null",
        assertThrows(NullPointerException.class, () -> RecipeCatalogs.entryFor(null)).getMessage());
    assertEquals(
        "id must not be null",
        assertThrows(NullPointerException.class, () -> RecipeCatalogs.lookupFor(null))
            .getMessage());
    assertEquals(
        "id must not be blank",
        assertThrows(IllegalArgumentException.class, () -> RecipeCatalogs.entryFor("   "))
            .getMessage());
    assertEquals(
        "id must not be blank",
        assertThrows(IllegalArgumentException.class, () -> RecipeCatalogs.lookupFor("   "))
            .getMessage());

    RecipeCatalogEntry assetBackedExample =
        new RecipeCatalogEntry(
            RecipeType.EXAMPLE,
            "ASSET_BACKED",
            "asset-backed-request.json",
            "summary",
            RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS,
            List.of("assets/example.xlsx"));

    assertEquals(
        List.of(assetBackedExample),
        RecipeCatalogValidation.copyRecipeCatalogEntries(List.of(assetBackedExample), "recipes"));

    assertEquals(
        "stepCount must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RecipeRequestProfile(
                        "NEW",
                        "NONE",
                        "FULL_XSSF",
                        "SUMMARY",
                        "DO_NOT_CALCULATE",
                        false,
                        -1,
                        List.of(),
                        List.of(),
                        List.of()))
            .getMessage());

    assertEquals(
        "type must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new RecipeCatalogEntry(
                        null,
                        "BROKEN_VIEW",
                        "broken-view.json",
                        "summary",
                        RecipeAdvisory.SELF_CONTAINED,
                        List.of()))
            .getMessage());

    IllegalArgumentException duplicateRecipes =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new RecipeCatalog(
                    ProtocolVersion.CURRENT, List.of(assetBackedExample, assetBackedExample)));
    assertEquals("recipes must not contain duplicate ASSET_BACKED", duplicateRecipes.getMessage());

    assertEquals(
        "SELF_CONTAINED recipes must not publish requiredWorkspacePaths",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RecipeCatalogEntry(
                        RecipeType.EXAMPLE,
                        "BROKEN_SELF_CONTAINED",
                        "broken-self-contained.json",
                        "summary",
                        RecipeAdvisory.SELF_CONTAINED,
                        List.of("assets/example.xlsx")))
            .getMessage());
    assertEquals(
        "REQUIRES_EXAMPLE_ASSETS recipes must publish requiredWorkspacePaths",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RecipeCatalogEntry(
                        RecipeType.EXAMPLE,
                        "BROKEN_ASSET_BACKED",
                        "broken-asset-backed.json",
                        "summary",
                        RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS,
                        List.of()))
            .getMessage());
    assertEquals(
        "discoveryTerms must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new TaskDiscoveryProfile(
                        List.of(),
                        new TaskIntentProfile(
                            List.of(TaskGoalKind.AUTHOR), List.of(TaskArtifactKind.WORKBOOK))))
            .getMessage());
    assertEquals(
        "SELF_CONTAINED examples must not publish requiredWorkspacePaths",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ShippedExampleEntry(
                        "BROKEN_EXAMPLE_SELF_CONTAINED",
                        "broken-example-self-contained.json",
                        "summary",
                        RecipeAdvisory.SELF_CONTAINED,
                        List.of("assets/example.xlsx")))
            .getMessage());
    assertEquals(
        "REQUIRES_EXAMPLE_ASSETS examples must publish requiredWorkspacePaths",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ShippedExampleEntry(
                        "BROKEN_EXAMPLE_ASSET_BACKED",
                        "broken-example-asset-backed.json",
                        "summary",
                        RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS,
                        List.of()))
            .getMessage());
  }

  @Test
  void discoveryRecordsRejectInvalidMandatoryShapesAndNormalizeOptionalDiagnosticFields() {
    assertEquals(
        "phases must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> new TaskWorkflow(List.of(), List.of("pitfall")))
            .getMessage());

    assertEquals(
        "discoveryTerms must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new TaskEntry(
                        "NO_TERMS",
                        List.of("office"),
                        new TaskDiscoveryProfile(
                            List.of(),
                            new TaskIntentProfile(
                                List.of(TaskGoalKind.AUTHOR), List.of(TaskArtifactKind.WORKBOOK))),
                        TaskTestFixtures.narrative("summary"),
                        profile(),
                        TaskTestFixtures.interactionProfile(),
                        TaskStarterContract.selfContained("no-terms-request.json"),
                        TaskTestFixtures.workflow(
                            List.of(
                                new TaskPhase(
                                    TaskPhasePurpose.AUTHOR,
                                    "Phase",
                                    "Objective",
                                    List.of(new TaskCapabilityRef("workbookSourceTypes", "NEW")),
                                    List.of("note"))))))
            .getMessage());

    assertEquals(
        "discoveryTerms must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new TaskDiscoveryProfile(
                        List.of(),
                        new TaskIntentProfile(
                            List.of(TaskGoalKind.AUTHOR), List.of(TaskArtifactKind.WORKBOOK))))
            .getMessage());

    assertEquals(
        "goals must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> new TaskIntentProfile(List.of(), List.of(TaskArtifactKind.WORKBOOK)))
            .getMessage());

    assertEquals(
        "artifacts must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> new TaskIntentProfile(List.of(TaskGoalKind.AUTHOR), List.of()))
            .getMessage());

    assertEquals(
        "problems must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new CommandError(
                        ProtocolVersion.CURRENT,
                        "print-recipe",
                        (List<GridGrindProblemDetail.Problem>) null))
            .getMessage());
    assertEquals(
        "problems must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CommandError(ProtocolVersion.CURRENT, "print-recipe", List.of()))
            .getMessage());
    assertEquals(
        "command must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> new CommandError(ProtocolVersion.CURRENT, "   ", List.of(problem())))
            .getMessage());
    assertEquals(
        "requestFileName must be one portable file name, not a repository path",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ShippedExampleEntry(
                        "BROKEN",
                        Path.of("examples", "budget-request.json").toString(),
                        "summary",
                        RecipeAdvisory.SELF_CONTAINED,
                        List.of()))
            .getMessage());
  }

  private static TaskExecutionProfile profile() {
    return TaskTestFixtures.profile();
  }

  private static GridGrindProblemDetail.Problem problem() {
    return GridGrindProblemDetail.Problem.of(
        GridGrindProblemCode.INVALID_ARGUMENTS,
        "message",
        new ProblemContext.ParseArguments(HostArgument.named("--lookup")));
  }
}
