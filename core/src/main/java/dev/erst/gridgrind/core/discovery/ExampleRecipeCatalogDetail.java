package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.metadata.Variant;
import java.util.List;
import java.util.Objects;

/** Detailed lookup payload for one built-in example recipe. */
@Variant(id = "EXAMPLE", summary = "Detailed lookup payload for one built-in example recipe.")
public record ExampleRecipeCatalogDetail(
    String id,
    String requestFileName,
    String summary,
    RecipeAdvisory advisory,
    List<String> requiredWorkspacePaths,
    List<String> intentTags,
    RecipeRequestProfile requestProfile)
    implements RecipeCatalogDetail {
  public ExampleRecipeCatalogDetail {
    id = DiscoveryValidation.requireNonBlank(id, "id");
    requestFileName = RecipeCatalogValidation.requirePortableRequestFileName(requestFileName);
    summary = DiscoveryValidation.requireNonBlank(summary, "summary");
    Objects.requireNonNull(advisory, "advisory must not be null");
    requiredWorkspacePaths =
        RecipeCatalogValidation.copyWorkspacePaths(
            requiredWorkspacePaths, "requiredWorkspacePaths");
    RecipeCatalogValidation.validateWorkspaceContract(advisory, requiredWorkspacePaths, "recipes");
    intentTags = DiscoveryValidation.copyStrings(intentTags, "intentTags");
    Objects.requireNonNull(requestProfile, "requestProfile must not be null");
  }
}
