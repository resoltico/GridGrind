package dev.erst.gridgrind.cli.command;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.discovery.Recipe;
import dev.erst.gridgrind.core.discovery.RecipeAdvisory;
import dev.erst.gridgrind.core.discovery.RecipeRegistry;
import dev.erst.gridgrind.core.discovery.ShippedExampleEntry;
import dev.erst.gridgrind.core.discovery.ShippedExamples;
import dev.erst.gridgrind.core.discovery.TaskCatalogs;
import dev.erst.gridgrind.core.discovery.TaskEntry;
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
        ShippedExamples.catalog().examples().stream().map(example -> example.id()).toList());
    publishedRecipeIds.addAll(
        TaskCatalogs.catalog().tasks().stream().map(task -> task.id()).toList());

    assertEquals(new LinkedHashSet<>(publishedRecipeIds).size(), publishedRecipeIds.size());
    assertTrue(
        publishedRecipeIds.stream().allMatch(id -> RecipeRegistry.recipeFor(id).isPresent()));
  }

  @Test
  void taskAndExampleViewsDerivePublishedFieldsFromRegistryDefinitions() {
    assertTrue(
        RecipeRegistry.recipeFor("WORKBOOK_HEALTH")
            .filter(recipe -> recipe.type() == dev.erst.gridgrind.core.discovery.RecipeType.EXAMPLE)
            .isPresent());
    assertTrue(
        RecipeRegistry.recipeFor("DASHBOARD")
            .filter(
                recipe ->
                    recipe.type() == dev.erst.gridgrind.core.discovery.RecipeType.TASK_STARTER)
            .isPresent());
    assertEquals(
        RecipeAdvisory.SELF_CONTAINED,
        RecipeRegistry.recipeFor("WORKBOOK_HEALTH").orElseThrow().advisory());
    assertEquals(
        List.of("package-security-assets/gridgrind-package-security.xlsx"),
        RecipeRegistry.recipeFor("PACKAGE_SECURITY_INSPECTION")
            .orElseThrow()
            .requiredWorkspacePaths());

    for (ShippedExampleEntry entry : ShippedExamples.catalog().examples()) {
      Recipe recipe = RecipeRegistry.recipeFor(entry.id()).orElseThrow();
      assertEquals(dev.erst.gridgrind.core.discovery.RecipeType.EXAMPLE, recipe.type());
      assertEquals(recipe.requestFileName(), entry.requestFileName());
      assertEquals(recipe.summary(), entry.summary());
      assertEquals(recipe.advisory(), entry.advisory());
      assertEquals(recipe.requiredWorkspacePaths(), entry.requiredWorkspacePaths());
    }

    for (TaskEntry entry : TaskCatalogs.catalog().tasks()) {
      Recipe recipe = RecipeRegistry.recipeFor(entry.id()).orElseThrow();
      assertEquals(dev.erst.gridgrind.core.discovery.RecipeType.TASK_STARTER, recipe.type());
      assertEquals(entry.narrative().summary(), recipe.summary());
      assertEquals(entry.intentTags(), recipe.intentTags());
      assertEquals(entry.starter().requestFileName(), recipe.requestFileName());
      assertEquals(entry.starter().advisory(), recipe.advisory());
      assertEquals(entry.starter().requiredWorkspacePaths(), recipe.requiredWorkspacePaths());
    }
  }
}
