package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.metadata.Union;

/** One detail view of a published recipe. */
@Union(discriminator = "type")
public sealed interface RecipeCatalogDetail extends RecipeCatalogDescriptor
    permits ExampleRecipeCatalogDetail, TaskStarterRecipeCatalogDetail {
  /** Published intent-tag vocabulary for this recipe, shared by both detail views. */
  java.util.List<String> intentTags();

  /** Exact runnable request profile derived from the packaged starter request. */
  RecipeRequestProfile requestProfile();
}
