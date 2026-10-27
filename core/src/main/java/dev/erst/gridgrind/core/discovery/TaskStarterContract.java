package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;

/** Portability contract for one printed published task starter request. */
public record TaskStarterContract(
    String requestFileName, RecipeAdvisory advisory, List<String> requiredWorkspacePaths) {
  public TaskStarterContract {
    requestFileName = TextValidation.requireNonBlank(requestFileName, "requestFileName");
    Objects.requireNonNull(advisory, "advisory must not be null");
    requiredWorkspacePaths =
        DiscoveryValidation.copyStringsAllowEmpty(requiredWorkspacePaths, "requiredWorkspacePaths");
    if (requestFileName.contains("/") || requestFileName.contains("\\")) {
      throw new IllegalArgumentException(
          "requestFileName must be one portable file name, not a repository path");
    }
    if (!requestFileName.endsWith(".json")) {
      throw new IllegalArgumentException("requestFileName must end with .json");
    }
    if (advisory == RecipeAdvisory.SELF_CONTAINED && !requiredWorkspacePaths.isEmpty()) {
      throw new IllegalArgumentException(
          "SELF_CONTAINED task starters must not publish requiredWorkspacePaths");
    }
    if (advisory == RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS && requiredWorkspacePaths.isEmpty()) {
      throw new IllegalArgumentException(
          "REQUIRES_EXAMPLE_ASSETS task starters must publish requiredWorkspacePaths");
    }
  }

  /** Creates one self-contained task starter contract with one portable request file name. */
  public static TaskStarterContract selfContained(String requestFileName) {
    return new TaskStarterContract(requestFileName, RecipeAdvisory.SELF_CONTAINED, List.of());
  }

  /** Creates one asset-backed task starter contract with workspace-relative asset requirements. */
  public static TaskStarterContract assetBacked(
      String requestFileName, String... requiredWorkspacePaths) {
    return new TaskStarterContract(
        requestFileName, RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS, List.of(requiredWorkspacePaths));
  }
}
