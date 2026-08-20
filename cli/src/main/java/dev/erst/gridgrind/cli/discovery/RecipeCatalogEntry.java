package dev.erst.gridgrind.cli.discovery;

import java.util.List;
import java.util.Objects;

/** One compact published recipe row in the unified machine-readable recipe catalog. */
public record RecipeCatalogEntry(
    RecipeView view,
    String id,
    String requestFileName,
    String summary,
    RecipeAdvisory advisory,
    List<String> requiredWorkspacePaths)
    implements RecipeCatalogDescriptor {
  public RecipeCatalogEntry {
    Objects.requireNonNull(view, "view must not be null");
    id = CliDiscoveryValidation.requireNonBlank(id, "id");
    requestFileName = CliRecipeCatalogValidation.requirePortableRequestFileName(requestFileName);
    summary = CliDiscoveryValidation.requireNonBlank(summary, "summary");
    Objects.requireNonNull(advisory, "advisory must not be null");
    requiredWorkspacePaths =
        CliRecipeCatalogValidation.copyWorkspacePaths(
            requiredWorkspacePaths, "requiredWorkspacePaths");
    CliRecipeCatalogValidation.validateWorkspaceContract(
        advisory, requiredWorkspacePaths, "recipes");
  }
}
