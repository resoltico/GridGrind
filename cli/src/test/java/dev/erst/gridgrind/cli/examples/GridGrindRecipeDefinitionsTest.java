package dev.erst.gridgrind.cli.examples;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.ExampleWorkspaceMode;
import dev.erst.gridgrind.cli.discovery.TaskEntry;
import dev.erst.gridgrind.cli.discovery.TaskStarterContract;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.GridGrindJsonOutput;
import java.util.List;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;

/** Coverage for the unified canonical recipe-definition source. */
class GridGrindRecipeDefinitionsTest {
  @Test
  void taskStarterContractRejectsIncompatibleWorkspaceModes() {
    IllegalArgumentException selfContainedFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskStarterContract(
                    "example.json",
                    ExampleWorkspaceMode.SELF_CONTAINED,
                    List.of("task-starter-assets/source.xlsx")));
    assertEquals(
        "SELF_CONTAINED task starters must not publish requiredWorkspacePaths",
        selfContainedFailure.getMessage());

    IllegalArgumentException assetBackedFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskStarterContract(
                    "example.json", ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS, List.of()));
    assertEquals(
        "REQUIRES_EXAMPLE_ASSETS task starters must publish requiredWorkspacePaths",
        assetBackedFailure.getMessage());
  }

  @Test
  void taskStarterContractRequiresPortableJsonFileNamesAndSupportsFactoryHelpers() {
    IllegalArgumentException repositoryPath =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskStarterContract(
                    "examples/example.json", ExampleWorkspaceMode.SELF_CONTAINED, List.of()));
    assertEquals(
        "requestFileName must be one portable file name, not a repository path",
        repositoryPath.getMessage());

    IllegalArgumentException windowsPath =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskStarterContract(
                    "examples\\example.json", ExampleWorkspaceMode.SELF_CONTAINED, List.of()));
    assertEquals(
        "requestFileName must be one portable file name, not a repository path",
        windowsPath.getMessage());

    IllegalArgumentException wrongSuffix =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskStarterContract(
                    "example.txt", ExampleWorkspaceMode.SELF_CONTAINED, List.of()));
    assertEquals("requestFileName must end with .json", wrongSuffix.getMessage());

    TaskStarterContract selfContained =
        TaskStarterContract.selfContained(
            TaskStarterRecipeSupport.taskRequestFileName("DASHBOARD"));
    assertEquals(ExampleWorkspaceMode.SELF_CONTAINED, selfContained.workspaceMode());
    assertEquals(List.of(), selfContained.requiredWorkspacePaths());

    TaskStarterContract assetBacked =
        TaskStarterContract.assetBacked(
            "example.json", "task-starter-assets/source.xlsx", "payloads/data.xml");
    assertEquals(ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS, assetBacked.workspaceMode());
    assertEquals(
        List.of("task-starter-assets/source.xlsx", "payloads/data.xml"),
        assetBacked.requiredWorkspacePaths());
  }

  @Test
  void registryPublishesStableTaskStarterPlansThroughTheSharedRecipeSurface() {
    assertTrue(GridGrindCliRecipeRegistry.recipeFor("NO_SUCH_TASK").isEmpty());

    WorkbookPlan request =
        GridGrindCliRecipeRegistry.recipeFor("DASHBOARD")
            .filter(
                recipe -> recipe.view() == dev.erst.gridgrind.cli.discovery.RecipeView.TASK_STARTER)
            .orElseThrow()
            .plan();
    assertEquals("NEW", textField(GridGrindJsonOutput.requestTree(request).path("source"), "type"));
    assertEquals("SAVE_AS", persistenceType(request));
    assertEquals("REPLACE", persistenceIfExists(request));
  }

  @Test
  void canonicalRecipeDefinitionsBackBothPublishedViews() {
    List<GridGrindRecipeDefinition> definitions = GridGrindRecipeDefinitions.definitions();

    assertEquals(
        definitions.stream().map(GridGrindRecipeDefinitionsTest::definitionId).distinct().count(),
        definitions.size());
    assertEquals(
        definitions.stream()
            .filter(GridGrindTaskRecipeDefinition.class::isInstance)
            .map(GridGrindTaskRecipeDefinition.class::cast)
            .map(GridGrindTaskRecipeDefinition::task)
            .toList(),
        GridGrindCliRecipeRegistry.taskEntries());
    assertEquals(
        definitions.stream()
            .filter(GridGrindExampleRecipeDefinition.class::isInstance)
            .map(GridGrindExampleRecipeDefinition.class::cast)
            .map(GridGrindExampleRecipeDefinition::id)
            .toList(),
        GridGrindShippedExamples.catalog().examples().stream().map(entry -> entry.id()).toList());

    for (GridGrindRecipeDefinition definition : definitions) {
      GridGrindCliRecipe definitionRecipe = publishedRecipe(definition);
      GridGrindCliRecipe publishedRecipe =
          GridGrindCliRecipeRegistry.recipeFor(definitionRecipe.id()).orElseThrow();
      assertEquals(definitionRecipe, publishedRecipe);

      if (definition instanceof GridGrindExampleRecipeDefinition example) {
        assertEquals(
            example.repositoryPlan(),
            GridGrindCliRecipeRegistry.repositoryExamplePlanFor(example.id()));
      } else if (definition instanceof GridGrindTaskRecipeDefinition task) {
        TaskEntry taskEntry =
            GridGrindCliRecipeRegistry.taskEntryFor(task.task().id()).orElseThrow();
        assertEquals(task.task(), taskEntry);
        assertEquals(task.starterPlan(), publishedRecipe.plan());
      } else {
        throw new AssertionError("Unknown recipe definition type " + definition.getClass());
      }
    }
  }

  private static String definitionId(GridGrindRecipeDefinition definition) {
    return publishedRecipe(definition).id();
  }

  private static GridGrindCliRecipe publishedRecipe(GridGrindRecipeDefinition definition) {
    if (definition instanceof GridGrindExampleRecipeDefinition example) {
      return example.recipe();
    }
    if (definition instanceof GridGrindTaskRecipeDefinition task) {
      return task.recipe();
    }
    throw new AssertionError("Unknown recipe definition type " + definition.getClass());
  }

  private static String persistenceType(WorkbookPlan request) {
    return textField(GridGrindJsonOutput.requestTree(request).path("persistence"), "type");
  }

  private static String persistenceIfExists(WorkbookPlan request) {
    return textField(GridGrindJsonOutput.requestTree(request).path("persistence"), "ifExists");
  }

  private static String textField(JsonNode node, String fieldName) {
    return java.util.Objects.requireNonNull(node.path(fieldName).stringValue());
  }
}
