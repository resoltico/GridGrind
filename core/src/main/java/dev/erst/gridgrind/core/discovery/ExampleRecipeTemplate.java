package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

/** Canonical published example recipe carrying discovery metadata plus both runnable plan views. */
record ExampleRecipeTemplate(
    String id,
    String requestFileName,
    String summary,
    RecipeAdvisory advisory,
    List<String> requiredWorkspacePaths,
    List<String> intentTags,
    WorkbookPlan builtInPlan,
    WorkbookPlan repositoryPlan)
    implements RecipeTemplate {
  ExampleRecipeTemplate {
    id = TextValidation.requireNonBlank(id, "id");
    if (!id.equals(id.toUpperCase(Locale.ROOT))) {
      throw new IllegalArgumentException("id must use upper-case discovery tokens");
    }
    requestFileName = TextValidation.requireNonBlank(requestFileName, "requestFileName");
    summary = TextValidation.requireNonBlank(summary, "summary");
    Objects.requireNonNull(advisory, "advisory must not be null");
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
    if (advisory == RecipeAdvisory.SELF_CONTAINED && !requiredWorkspacePaths.isEmpty()) {
      throw new IllegalArgumentException(
          "SELF_CONTAINED example recipes must not publish requiredWorkspacePaths");
    }
    if (advisory == RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS && requiredWorkspacePaths.isEmpty()) {
      throw new IllegalArgumentException(
          "REQUIRES_EXAMPLE_ASSETS example recipes must publish requiredWorkspacePaths");
    }
  }

  Recipe recipe() {
    return new Recipe(
        id,
        RecipeType.EXAMPLE,
        requestFileName,
        summary,
        advisory,
        requiredWorkspacePaths,
        intentTags,
        builtInPlan);
  }

  private static List<String> copyNonBlankList(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<String> copied =
        values.stream().map(value -> TextValidation.requireNonBlank(value, fieldName)).toList();
    if (copied.isEmpty()) {
      throw new IllegalArgumentException(fieldName + " must not be empty");
    }
    return copied;
  }
}
