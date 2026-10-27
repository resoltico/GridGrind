package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** Public metadata for one generated built-in example workbook plan. */
public record ShippedExampleEntry(
    String id,
    String requestFileName,
    String summary,
    RecipeAdvisory advisory,
    java.util.List<String> requiredWorkspacePaths) {
  public ShippedExampleEntry {
    id = TextValidation.requireNonBlank(id, "id");
    requestFileName = RecipeCatalogValidation.requirePortableRequestFileName(requestFileName);
    summary = TextValidation.requireNonBlank(summary, "summary");
    java.util.Objects.requireNonNull(advisory, "advisory must not be null");
    requiredWorkspacePaths =
        RecipeCatalogValidation.copyWorkspacePaths(
            requiredWorkspacePaths, "requiredWorkspacePaths");
    RecipeCatalogValidation.validateWorkspaceContract(advisory, requiredWorkspacePaths, "examples");
  }
}
