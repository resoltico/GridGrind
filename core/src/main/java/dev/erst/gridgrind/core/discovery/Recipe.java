package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

/** Public recipe type published for discovery for built-in examples and task starters. */
public record Recipe(
    String id,
    RecipeType type,
    String requestFileName,
    String summary,
    RecipeAdvisory advisory,
    List<String> requiredWorkspacePaths,
    List<String> intentTags,
    WorkbookPlan plan) {
  public Recipe {
    id = TextValidation.requireNonBlank(id, "id");
    if (!id.equals(id.toUpperCase(Locale.ROOT))) {
      throw new IllegalArgumentException("id must use upper-case discovery tokens");
    }
    Objects.requireNonNull(type, "type must not be null");
    requestFileName = TextValidation.requireNonBlank(requestFileName, "requestFileName");
    summary = TextValidation.requireNonBlank(summary, "summary");
    Objects.requireNonNull(advisory, "advisory must not be null");
    requiredWorkspacePaths =
        List.copyOf(
            Objects.requireNonNull(
                requiredWorkspacePaths, "requiredWorkspacePaths must not be null"));
    intentTags = copyNonBlankList(intentTags, "intentTags");
    Objects.requireNonNull(plan, "plan must not be null");
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
