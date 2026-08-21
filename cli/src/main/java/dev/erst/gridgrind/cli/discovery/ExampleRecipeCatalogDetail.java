package dev.erst.gridgrind.cli.discovery;

import java.util.List;
import java.util.Objects;

/** Detailed lookup payload for one built-in example recipe. */
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
    id = CliDiscoveryValidation.requireNonBlank(id, "id");
    requestFileName = CliRecipeCatalogValidation.requirePortableRequestFileName(requestFileName);
    summary = CliDiscoveryValidation.requireNonBlank(summary, "summary");
    Objects.requireNonNull(advisory, "advisory must not be null");
    requiredWorkspacePaths =
        CliRecipeCatalogValidation.copyWorkspacePaths(
            requiredWorkspacePaths, "requiredWorkspacePaths");
    CliRecipeCatalogValidation.validateWorkspaceContract(
        advisory, requiredWorkspacePaths, "recipes");
    intentTags = CliDiscoveryValidation.copyStrings(intentTags, "intentTags");
    Objects.requireNonNull(requestProfile, "requestProfile must not be null");
  }
}
