package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.cli.discovery.ExampleWorkspaceMode;
import dev.erst.gridgrind.cli.discovery.RecipeView;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

/** Canonical published example recipe carrying discovery metadata plus both runnable plan views. */
record GridGrindExampleRecipeDefinition(
    String id,
    String requestFileName,
    String summary,
    ExampleWorkspaceMode workspaceMode,
    List<String> requiredWorkspacePaths,
    List<String> intentTags,
    WorkbookPlan builtInPlan,
    WorkbookPlan repositoryPlan)
    implements GridGrindRecipeDefinition {
  GridGrindExampleRecipeDefinition {
    id = requireNonBlank(id, "id");
    if (!id.equals(id.toUpperCase(Locale.ROOT))) {
      throw new IllegalArgumentException("id must use upper-case discovery tokens");
    }
    requestFileName = requireNonBlank(requestFileName, "requestFileName");
    summary = requireNonBlank(summary, "summary");
    Objects.requireNonNull(workspaceMode, "workspaceMode must not be null");
    requiredWorkspacePaths = List.copyOf(requiredWorkspacePaths);
    intentTags = copyNonBlankList(intentTags, "intentTags");
    Objects.requireNonNull(builtInPlan, "builtInPlan must not be null");
    Objects.requireNonNull(repositoryPlan, "repositoryPlan must not be null");
    if (requestFileName.contains("/") || requestFileName.contains("\\")) {
      throw new IllegalArgumentException(
          "requestFileName must be one portable file name, not a repository path");
    }
    if (!requestFileName.endsWith(".json")) {
      throw new IllegalArgumentException("requestFileName must end with .json");
    }
    if (workspaceMode == ExampleWorkspaceMode.SELF_CONTAINED && !requiredWorkspacePaths.isEmpty()) {
      throw new IllegalArgumentException(
          "SELF_CONTAINED example recipes must not publish requiredWorkspacePaths");
    }
    if (workspaceMode == ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS
        && requiredWorkspacePaths.isEmpty()) {
      throw new IllegalArgumentException(
          "REQUIRES_EXAMPLE_ASSETS example recipes must publish requiredWorkspacePaths");
    }
  }

  GridGrindCliRecipe recipe() {
    return new GridGrindCliRecipe(
        id,
        RecipeView.EXAMPLE,
        requestFileName,
        summary,
        workspaceMode,
        requiredWorkspacePaths,
        intentTags,
        builtInPlan);
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static List<String> copyNonBlankList(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<String> copied = values.stream().map(value -> requireNonBlank(value, fieldName)).toList();
    if (copied.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    return copied;
  }
}
