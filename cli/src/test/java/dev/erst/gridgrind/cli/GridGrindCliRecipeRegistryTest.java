package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.cli.discovery.ExampleWorkspaceMode;
import dev.erst.gridgrind.cli.discovery.GridGrindTaskCatalog;
import dev.erst.gridgrind.cli.discovery.ShippedExampleEntry;
import dev.erst.gridgrind.cli.discovery.TaskEntry;
import dev.erst.gridgrind.cli.examples.GridGrindCliRecipe;
import dev.erst.gridgrind.cli.examples.GridGrindCliRecipeRegistry;
import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Regression coverage for the unified CLI recipe registry. */
class GridGrindCliRecipeRegistryTest {
  @Test
  void registryMaintainsGloballyUniqueRecipeIdsAndStableViews() {
    List<String> publishedRecipeIds = new ArrayList<>();
    publishedRecipeIds.addAll(
        GridGrindShippedExamples.catalog().examples().stream()
            .map(example -> example.id())
            .toList());
    publishedRecipeIds.addAll(
        GridGrindTaskCatalog.catalog().tasks().stream().map(task -> task.id()).toList());

    assertEquals(new LinkedHashSet<>(publishedRecipeIds).size(), publishedRecipeIds.size());
    assertTrue(
        publishedRecipeIds.stream()
            .allMatch(id -> GridGrindCliRecipeRegistry.recipeFor(id).isPresent()));
  }

  @Test
  void taskAndExampleViewsDerivePublishedFieldsFromRegistryDefinitions() {
    assertTrue(
        GridGrindCliRecipeRegistry.recipeFor("WORKBOOK_HEALTH")
            .filter(recipe -> recipe.view() == dev.erst.gridgrind.cli.discovery.RecipeView.EXAMPLE)
            .isPresent());
    assertTrue(
        GridGrindCliRecipeRegistry.recipeFor("DASHBOARD")
            .filter(
                recipe -> recipe.view() == dev.erst.gridgrind.cli.discovery.RecipeView.TASK_STARTER)
            .isPresent());
    assertEquals(
        ExampleWorkspaceMode.SELF_CONTAINED,
        GridGrindCliRecipeRegistry.recipeFor("WORKBOOK_HEALTH").orElseThrow().workspaceMode());
    assertEquals(
        List.of("package-security-assets/gridgrind-package-security.xlsx"),
        GridGrindCliRecipeRegistry.recipeFor("PACKAGE_SECURITY_INSPECTION")
            .orElseThrow()
            .requiredWorkspacePaths());

    for (ShippedExampleEntry entry : GridGrindShippedExamples.catalog().examples()) {
      GridGrindCliRecipe recipe = GridGrindCliRecipeRegistry.recipeFor(entry.id()).orElseThrow();
      assertEquals(dev.erst.gridgrind.cli.discovery.RecipeView.EXAMPLE, recipe.view());
      assertEquals(recipe.requestFileName(), entry.requestFileName());
      assertEquals(recipe.summary(), entry.summary());
      assertEquals(recipe.workspaceMode(), entry.workspaceMode());
      assertEquals(recipe.requiredWorkspacePaths(), entry.requiredWorkspacePaths());
    }

    for (TaskEntry entry : GridGrindTaskCatalog.catalog().tasks()) {
      GridGrindCliRecipe recipe = GridGrindCliRecipeRegistry.recipeFor(entry.id()).orElseThrow();
      assertEquals(dev.erst.gridgrind.cli.discovery.RecipeView.TASK_STARTER, recipe.view());
      assertEquals(entry.narrative().summary(), recipe.summary());
      assertEquals(entry.intentTags(), recipe.intentTags());
      assertEquals(entry.starter().requestFileName(), recipe.requestFileName());
      assertEquals(entry.starter().workspaceMode(), recipe.workspaceMode());
      assertEquals(entry.starter().requiredWorkspacePaths(), recipe.requiredWorkspacePaths());
    }
  }
}
