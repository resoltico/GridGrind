package dev.erst.gridgrind.cli.discovery;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.examples.GridGrindCliRecipeRegistry;
import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.contract.catalog.TypeEntry;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.CliArgument;
import java.nio.file.Path;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct coverage for CLI discovery validation helpers and task-definition accessors. */
class CliDiscoveryValidationCoverageTest {
  @Test
  void validatesPrimitiveCopyHelpersAndNullGuards() {
    GridGrindProtocolVersion protocolVersion = GridGrindProtocolVersion.current();
    TypeEntry typeEntry = GridGrindProtocolCatalog.catalog().plainTypes().getFirst().type();
    ShippedExampleEntry exampleEntry = GridGrindShippedExamples.catalog().examples().getFirst();
    TaskEntry taskEntry = GridGrindTaskCatalog.catalog().tasks().getFirst();
    TaskPhase taskPhase = taskEntry.workflow().phases().getFirst();
    TaskCapabilityRef capabilityRef = taskPhase.capabilityRefs().getFirst();

    assertEquals(protocolVersion, CliDiscoveryValidation.requireProtocolVersion(protocolVersion));
    assertEquals(
        List.of("alpha", "beta"),
        CliDiscoveryValidation.copyStrings(List.of("alpha", "beta"), "values"));
    assertEquals(
        List.of(typeEntry), CliDiscoveryValidation.copyTypeEntries(List.of(typeEntry), "entries"));
    assertEquals(
        List.of(taskEntry), CliDiscoveryValidation.copyTaskEntries(List.of(taskEntry), "tasks"));
    assertEquals(
        List.of(taskPhase), CliDiscoveryValidation.copyTaskPhases(List.of(taskPhase), "phases"));
    assertEquals(
        List.of(capabilityRef),
        CliDiscoveryValidation.copyTaskCapabilityRefs(List.of(capabilityRef), "capabilityRefs"));
    assertEquals(
        List.of(exampleEntry),
        CliDiscoveryValidation.copyExampleEntries(List.of(exampleEntry), "examples"));
    assertEquals(
        Optional.of("value"),
        CliDiscoveryValidation.normalizeOptionalString(Optional.of("value"), "label"));
    assertEquals(
        Optional.empty(),
        CliDiscoveryValidation.normalizeOptionalString(Optional.empty(), "label"));
    assertEquals(
        Optional.of("value"),
        CliDiscoveryValidation.copyOptionalArbitraryString(Optional.of("value"), "query"));
    assertEquals(
        Optional.empty(),
        CliDiscoveryValidation.copyOptionalArbitraryString(Optional.empty(), "query"));

    assertEquals(
        "protocolVersion must not be null",
        assertThrows(
                NullPointerException.class,
                () -> CliDiscoveryValidation.requireProtocolVersion(null))
            .getMessage());
    assertEquals(
        "entries must not be null",
        assertThrows(
                NullPointerException.class,
                () -> CliDiscoveryValidation.copyTypeEntries(null, "entries"))
            .getMessage());
    assertEquals(
        "label must not be null",
        assertThrows(
                NullPointerException.class,
                () -> CliDiscoveryValidation.normalizeOptionalString(null, "label"))
            .getMessage());
    assertEquals(
        "label must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> CliDiscoveryValidation.normalizeOptionalString(Optional.of("   "), "label"))
            .getMessage());
    assertEquals(
        "query must not be null",
        assertThrows(
                NullPointerException.class,
                () -> CliDiscoveryValidation.copyOptionalArbitraryString(null, "query"))
            .getMessage());
    assertEquals(
        "entries must not contain nulls",
        assertThrows(
                NullPointerException.class,
                () ->
                    CliDiscoveryValidation.copyTypeEntries(
                        java.util.Arrays.asList((TypeEntry) null), "entries"))
            .getMessage());
  }

  @Test
  void taskCatalogDerivesCanonicalEntriesFromTheSharedRegistryAndRejectsInvalidLookupIds() {
    TaskCatalog catalog = GridGrindTaskCatalog.catalog();
    List<TaskEntry> entries = GridGrindCliRecipeRegistry.taskEntries();

    assertFalse(entries.isEmpty());
    assertNotSame(entries, catalog.tasks());
    assertEquals(entries, catalog.tasks());
    assertTrue(GridGrindTaskCatalog.entryFor("DASHBOARD").isPresent());
    assertTrue(GridGrindTaskCatalog.entryFor("WORKBOOK_MAINTENANCE").isPresent());
    assertTrue(GridGrindTaskCatalog.entryFor("NO_SUCH_TASK").isEmpty());
    GridGrindTaskCatalog.validateCapabilityReferences(catalog);

    assertEquals(
        "id must not be null",
        assertThrows(NullPointerException.class, () -> GridGrindTaskCatalog.entryFor(null))
            .getMessage());
    assertEquals(
        "id must not be blank",
        assertThrows(IllegalArgumentException.class, () -> GridGrindTaskCatalog.entryFor("   "))
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
                    GridGrindTaskCatalog.validateCapabilityReferences(
                        new TaskCatalog(GridGrindProtocolVersion.current(), List.of(brokenTask))))
            .getMessage());
  }

  @Test
  void recipeCatalogValidationRejectsDuplicateIdsAndInvalidWorkspaceContracts() {
    assertTrue(GridGrindRecipeCatalog.entryFor("DASHBOARD").isPresent());
    assertTrue(GridGrindRecipeCatalog.lookupFor("DASHBOARD").isPresent());
    assertEquals(
        "id must not be null",
        assertThrows(NullPointerException.class, () -> GridGrindRecipeCatalog.entryFor(null))
            .getMessage());
    assertEquals(
        "id must not be null",
        assertThrows(NullPointerException.class, () -> GridGrindRecipeCatalog.lookupFor(null))
            .getMessage());
    assertEquals(
        "id must not be blank",
        assertThrows(IllegalArgumentException.class, () -> GridGrindRecipeCatalog.entryFor("   "))
            .getMessage());
    assertEquals(
        "id must not be blank",
        assertThrows(IllegalArgumentException.class, () -> GridGrindRecipeCatalog.lookupFor("   "))
            .getMessage());

    RecipeCatalogEntry assetBackedExample =
        new RecipeCatalogEntry(
            RecipeView.EXAMPLE,
            "ASSET_BACKED",
            "asset-backed-request.json",
            "summary",
            ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS,
            List.of("assets/example.xlsx"));

    assertEquals(
        List.of(assetBackedExample),
        CliRecipeCatalogValidation.copyRecipeCatalogEntries(
            List.of(assetBackedExample), "recipes"));

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
        "view must not be null",
        assertThrows(
                NullPointerException.class,
                () ->
                    new RecipeCatalogEntry(
                        null,
                        "BROKEN_VIEW",
                        "broken-view.json",
                        "summary",
                        ExampleWorkspaceMode.SELF_CONTAINED,
                        List.of()))
            .getMessage());

    IllegalArgumentException duplicateRecipes =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new RecipeCatalog(
                    GridGrindProtocolVersion.current(),
                    List.of(assetBackedExample, assetBackedExample)));
    assertEquals("recipes must not contain duplicate ASSET_BACKED", duplicateRecipes.getMessage());

    assertEquals(
        "SELF_CONTAINED recipes must not publish requiredWorkspacePaths",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RecipeCatalogEntry(
                        RecipeView.EXAMPLE,
                        "BROKEN_SELF_CONTAINED",
                        "broken-self-contained.json",
                        "summary",
                        ExampleWorkspaceMode.SELF_CONTAINED,
                        List.of("assets/example.xlsx")))
            .getMessage());
    assertEquals(
        "REQUIRES_EXAMPLE_ASSETS recipes must publish requiredWorkspacePaths",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new RecipeCatalogEntry(
                        RecipeView.EXAMPLE,
                        "BROKEN_ASSET_BACKED",
                        "broken-asset-backed.json",
                        "summary",
                        ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS,
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
                        ExampleWorkspaceMode.SELF_CONTAINED,
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
                        ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS,
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
                                    List.of(new TaskCapabilityRef("sourceTypes", "NEW")),
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
                        GridGrindProtocolVersion.current(),
                        "print-recipe",
                        (List<GridGrindProblemDetail.Problem>) null))
            .getMessage());
    assertEquals(
        "problems must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new CommandError(
                        GridGrindProtocolVersion.current(),
                        "print-recipe",
                        List.of()))
            .getMessage());
    assertEquals(
        "command must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new CommandError(
                        GridGrindProtocolVersion.current(),
                        "   ",
                        List.of(problem())))
            .getMessage());
    assertEquals(
        "responsePath must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new CliTransportNotice(
                        CliTransportNotice.Destination.STDOUT, Optional.of("   ")))
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
                        ExampleWorkspaceMode.SELF_CONTAINED,
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
        new ProblemContext.ParseArguments(CliArgument.named("--lookup")));
  }
}
