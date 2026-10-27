package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import java.util.Objects;

/** Canonical published task recipe carrying both discovery metadata and the starter plan. */
record TaskRecipeTemplate(TaskEntry task, WorkbookPlan starterPlan) implements RecipeTemplate {
  TaskRecipeTemplate {
    Objects.requireNonNull(task, "task must not be null");
    Objects.requireNonNull(starterPlan, "starterPlan must not be null");
  }

  Recipe recipe() {
    return new Recipe(
        task.id(),
        RecipeType.TASK_STARTER,
        task.starter().requestFileName(),
        task.narrative().summary(),
        task.starter().advisory(),
        task.starter().requiredWorkspacePaths(),
        task.intentTags(),
        starterPlan);
  }
}
