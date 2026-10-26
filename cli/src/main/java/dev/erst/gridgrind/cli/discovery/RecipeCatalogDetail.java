package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.core.metadata.Union;

/** One view-specific detail payload returned by {@code --print-recipe-catalog --lookup <id>}. */
@Union(discriminator = "type")
public sealed interface RecipeCatalogDetail extends RecipeCatalogDescriptor
    permits ExampleRecipeCatalogDetail, TaskStarterRecipeCatalogDetail {
  /** Published intent-tag vocabulary for this recipe, shared by both detail views. */
  java.util.List<String> intentTags();

  /** Exact runnable request profile derived from the packaged starter request. */
  RecipeRequestProfile requestProfile();
}
