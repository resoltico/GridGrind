package dev.erst.gridgrind.core.discovery;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Coverage for the unified canonical recipe-definition source. */
class GridGrindRecipeDefinitionsTest {
  @Test
  void taskStarterContractRejectsIncompatibleAdvisories() {
    IllegalArgumentException selfContainedFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskStarterContract(
                    "example.json",
                    RecipeAdvisory.SELF_CONTAINED,
                    List.of("task-starter-assets/source.xlsx")));
    assertEquals(
        "SELF_CONTAINED task starters must not publish requiredWorkspacePaths",
        selfContainedFailure.getMessage());

    IllegalArgumentException assetBackedFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskStarterContract(
                    "example.json", RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS, List.of()));
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
                    "examples/example.json", RecipeAdvisory.SELF_CONTAINED, List.of()));
    assertEquals(
        "requestFileName must be one portable file name, not a repository path",
        repositoryPath.getMessage());

    IllegalArgumentException windowsPath =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new TaskStarterContract(
                    "examples\\example.json", RecipeAdvisory.SELF_CONTAINED, List.of()));
    assertEquals(
        "requestFileName must be one portable file name, not a repository path",
        windowsPath.getMessage());

    IllegalArgumentException wrongSuffix =
        assertThrows(
            IllegalArgumentException.class,
            () -> new TaskStarterContract("example.txt", RecipeAdvisory.SELF_CONTAINED, List.of()));
    assertEquals("requestFileName must end with .json", wrongSuffix.getMessage());

    TaskStarterContract selfContained =
        TaskStarterContract.selfContained(TaskStarters.taskRequestFileName("DASHBOARD"));
    assertEquals(RecipeAdvisory.SELF_CONTAINED, selfContained.advisory());
    assertEquals(List.of(), selfContained.requiredWorkspacePaths());

    TaskStarterContract assetBacked =
        TaskStarterContract.assetBacked(
            "example.json", "task-starter-assets/source.xlsx", "payloads/data.xml");
    assertEquals(RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS, assetBacked.advisory());
    assertEquals(
        List.of("task-starter-assets/source.xlsx", "payloads/data.xml"),
        assetBacked.requiredWorkspacePaths());
  }

  @Test
  void registryPublishesStableTaskStarterPlansThroughTheSharedRecipeSurface() {
    assertTrue(RecipeRegistry.recipeFor("NO_SUCH_TASK").isEmpty());

    WorkbookPlan request =
        RecipeRegistry.recipeFor("DASHBOARD")
            .filter(
                recipe ->
                    recipe.type() == dev.erst.gridgrind.core.discovery.RecipeType.TASK_STARTER)
            .orElseThrow()
            .plan();
    org.junit.jupiter.api.Assertions.assertInstanceOf(
        WorkbookPlan.WorkbookSource.New.class, request.source());
    var persistence =
        org.junit.jupiter.api.Assertions.assertInstanceOf(
            WorkbookPlan.WorkbookPersistence.SaveAs.class, request.persistence());
    assertEquals(WorkbookPlan.WorkbookPersistence.IfExists.REPLACE, persistence.ifExists());
  }

  @Test
  void canonicalRecipeDefinitionsBackBothPublishedViews() {
    List<RecipeTemplate> definitions = RecipeTemplates.definitions();

    assertEquals(
        definitions.stream().map(GridGrindRecipeDefinitionsTest::definitionId).distinct().count(),
        definitions.size());
    assertEquals(
        definitions.stream()
            .filter(TaskRecipeTemplate.class::isInstance)
            .map(TaskRecipeTemplate.class::cast)
            .map(TaskRecipeTemplate::task)
            .toList(),
        RecipeRegistry.taskEntries());
    assertEquals(
        definitions.stream()
            .filter(ExampleRecipeTemplate.class::isInstance)
            .map(ExampleRecipeTemplate.class::cast)
            .map(ExampleRecipeTemplate::id)
            .toList(),
        ShippedExamples.catalog().examples().stream().map(entry -> entry.id()).toList());

    for (RecipeTemplate definition : definitions) {
      Recipe definitionRecipe = publishedRecipe(definition);
      Recipe publishedRecipe = RecipeRegistry.recipeFor(definitionRecipe.id()).orElseThrow();
      assertEquals(definitionRecipe, publishedRecipe);

      if (definition instanceof ExampleRecipeTemplate example) {
        assertEquals(
            example.repositoryPlan(), RecipeRegistry.repositoryExamplePlanFor(example.id()));
      } else if (definition instanceof TaskRecipeTemplate task) {
        TaskEntry taskEntry = RecipeRegistry.taskEntryFor(task.task().id()).orElseThrow();
        assertEquals(task.task(), taskEntry);
        assertEquals(task.starterPlan(), publishedRecipe.plan());
      } else {
        throw new AssertionError("Unknown recipe definition type " + definition.getClass());
      }
    }
  }

  private static String definitionId(RecipeTemplate definition) {
    return publishedRecipe(definition).id();
  }

  private static Recipe publishedRecipe(RecipeTemplate definition) {
    if (definition instanceof ExampleRecipeTemplate example) {
      return example.recipe();
    }
    if (definition instanceof TaskRecipeTemplate task) {
      return task.recipe();
    }
    throw new AssertionError("Unknown recipe definition type " + definition.getClass());
  }
}
