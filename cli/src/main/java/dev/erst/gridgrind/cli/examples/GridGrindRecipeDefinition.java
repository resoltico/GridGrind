package dev.erst.gridgrind.cli.examples;

/** Canonical published recipe marker shared by example and task-starter variants. */
sealed interface GridGrindRecipeDefinition
    permits GridGrindExampleRecipeDefinition, GridGrindTaskRecipeDefinition {}
