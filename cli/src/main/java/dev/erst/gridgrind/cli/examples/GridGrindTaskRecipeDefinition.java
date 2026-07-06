package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.cli.discovery.RecipeView;
import dev.erst.gridgrind.cli.discovery.TaskEntry;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.Objects;

/** Canonical published task recipe carrying both discovery metadata and the starter plan. */
record GridGrindTaskRecipeDefinition(TaskEntry task, WorkbookPlan starterPlan)
    implements GridGrindRecipeDefinition {
  GridGrindTaskRecipeDefinition {
    Objects.requireNonNull(task, "task must not be null");
    Objects.requireNonNull(starterPlan, "starterPlan must not be null");
  }

  GridGrindCliRecipe recipe() {
    return new GridGrindCliRecipe(
        task.id(),
        RecipeView.TASK_STARTER,
        task.starter().requestFileName(),
        task.narrative().summary(),
        task.starter().workspaceMode(),
        task.starter().requiredWorkspacePaths(),
        task.intentTags(),
        starterPlan);
  }
}
