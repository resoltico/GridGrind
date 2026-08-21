package dev.erst.gridgrind.cli.discovery;

import java.util.List;
import java.util.Objects;

/** Recipe-catalog-specific validation for portable file names and workspace contracts. */
final class CliRecipeCatalogValidation {
  private CliRecipeCatalogValidation() {}

  static List<RecipeCatalogEntry> copyRecipeCatalogEntries(
      List<RecipeCatalogEntry> entries, String fieldName) {
    return CliDiscoveryValidation.copyUnique(entries, fieldName, RecipeCatalogEntry::id);
  }

  static List<String> copyWorkspacePaths(List<String> values, String fieldName) {
    return CliDiscoveryValidation.copyStringsAllowEmpty(values, fieldName);
  }

  static String requirePortableRequestFileName(String requestFileName) {
    String value = CliDiscoveryValidation.requireNonBlank(requestFileName, "requestFileName");
    if (value.contains("/") || value.contains("\\")) {
      throw new IllegalArgumentException(
          "requestFileName must be one portable file name, not a repository path");
    }
    if (!value.endsWith(".json")) {
      throw new IllegalArgumentException("requestFileName must end with .json");
    }
    return value;
  }

  static void validateWorkspaceContract(
      RecipeAdvisory advisory, List<String> requiredWorkspacePaths, String noun) {
    Objects.requireNonNull(advisory, "advisory must not be null");
    Objects.requireNonNull(requiredWorkspacePaths, "requiredWorkspacePaths must not be null");
    if (advisory == RecipeAdvisory.SELF_CONTAINED && !requiredWorkspacePaths.isEmpty()) {
      throw new IllegalArgumentException(
          "SELF_CONTAINED " + noun + " must not publish requiredWorkspacePaths");
    }
    if (advisory == RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS && requiredWorkspacePaths.isEmpty()) {
      throw new IllegalArgumentException(
          "REQUIRES_EXAMPLE_ASSETS " + noun + " must publish requiredWorkspacePaths");
    }
  }
}
