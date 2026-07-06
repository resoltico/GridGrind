package dev.erst.gridgrind.cli.discovery;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** One view-specific detail payload returned by {@code --print-recipe-catalog --lookup <id>}. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "view")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ExampleRecipeCatalogDetail.class, name = "EXAMPLE"),
  @JsonSubTypes.Type(value = TaskStarterRecipeCatalogDetail.class, name = "TASK_STARTER")
})
public sealed interface RecipeCatalogDetail extends RecipeCatalogDescriptor
    permits ExampleRecipeCatalogDetail, TaskStarterRecipeCatalogDetail {
  /** Published intent-tag vocabulary for this recipe, shared by both detail views. */
  java.util.List<String> intentTags();

  /** Exact runnable request profile derived from the packaged starter request. */
  RecipeRequestProfile requestProfile();
}
