package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.cli.discovery.RecipeAdvisory;
import dev.erst.gridgrind.cli.discovery.RecipeView;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

/** Public recipe view published by the CLI for built-in examples and task starters. */
public record GridGrindCliRecipe(
    String id,
    RecipeView view,
    String requestFileName,
    String summary,
    RecipeAdvisory advisory,
    List<String> requiredWorkspacePaths,
    List<String> intentTags,
    WorkbookPlan plan) {
  public GridGrindCliRecipe {
    id = requireNonBlank(id, "id");
    if (!id.equals(id.toUpperCase(Locale.ROOT))) {
      throw new IllegalArgumentException("id must use upper-case discovery tokens");
    }
    Objects.requireNonNull(view, "view must not be null");
    requestFileName = requireNonBlank(requestFileName, "requestFileName");
    summary = requireNonBlank(summary, "summary");
    Objects.requireNonNull(advisory, "advisory must not be null");
    requiredWorkspacePaths =
        List.copyOf(
            Objects.requireNonNull(
                requiredWorkspacePaths, "requiredWorkspacePaths must not be null"));
    intentTags = copyNonBlankList(intentTags, "intentTags");
    Objects.requireNonNull(plan, "plan must not be null");
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
