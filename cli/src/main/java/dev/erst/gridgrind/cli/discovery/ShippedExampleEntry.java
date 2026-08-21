package dev.erst.gridgrind.cli.discovery;

/** Public metadata for one generated built-in example workbook plan. */
public record ShippedExampleEntry(
    String id,
    String requestFileName,
    String summary,
    RecipeAdvisory advisory,
    java.util.List<String> requiredWorkspacePaths) {
  public ShippedExampleEntry {
    id = CliDiscoveryValidation.requireNonBlank(id, "id");
    requestFileName = CliRecipeCatalogValidation.requirePortableRequestFileName(requestFileName);
    summary = CliDiscoveryValidation.requireNonBlank(summary, "summary");
    java.util.Objects.requireNonNull(advisory, "advisory must not be null");
    requiredWorkspacePaths =
        CliRecipeCatalogValidation.copyWorkspacePaths(
            requiredWorkspacePaths, "requiredWorkspacePaths");
    CliRecipeCatalogValidation.validateWorkspaceContract(
        advisory, requiredWorkspacePaths, "examples");
  }
}
