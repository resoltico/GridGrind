package dev.erst.gridgrind.core.discovery;

/** Canonical published recipe marker shared by example and task-starter variants. */
sealed interface RecipeTemplate permits ExampleRecipeTemplate, TaskRecipeTemplate {}
