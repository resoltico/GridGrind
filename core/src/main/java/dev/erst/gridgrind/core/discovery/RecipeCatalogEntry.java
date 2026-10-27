package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;

/** One compact published recipe row in the unified machine-readable recipe catalog. */
public record RecipeCatalogEntry(
    RecipeType type,
    String id,
    String requestFileName,
    String summary,
    RecipeAdvisory advisory,
    List<String> requiredWorkspacePaths)
    implements RecipeCatalogDescriptor {
  public RecipeCatalogEntry {
    Objects.requireNonNull(type, "type must not be null");
    id = TextValidation.requireNonBlank(id, "id");
    requestFileName = RecipeCatalogValidation.requirePortableRequestFileName(requestFileName);
    summary = TextValidation.requireNonBlank(summary, "summary");
    Objects.requireNonNull(advisory, "advisory must not be null");
    requiredWorkspacePaths =
        RecipeCatalogValidation.copyWorkspacePaths(
            requiredWorkspacePaths, "requiredWorkspacePaths");
    RecipeCatalogValidation.validateWorkspaceContract(advisory, requiredWorkspacePaths, "recipes");
  }
}
