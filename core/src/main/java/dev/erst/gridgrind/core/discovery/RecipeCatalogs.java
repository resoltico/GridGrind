package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Optional;
import java.util.Set;

/** Unified public recipe catalog layered on top of the canonical recipe registry. */
public final class RecipeCatalogs {
  private RecipeCatalogs() {}

  /** Returns the full built-in recipe catalog across examples and task starters. */
  public static RecipeCatalog catalog() {
    return new RecipeCatalog(ProtocolVersion.CURRENT, recipeEntries());
  }

  /** Returns one recipe catalog entry by stable id, or empty when unknown. */
  public static Optional<RecipeCatalogEntry> entryFor(String id) {
    java.util.Objects.requireNonNull(id, "id must not be null");
    String lookup = id.trim();
    if (lookup.isBlank()) {
      throw new IllegalArgumentException("id must not be blank");
    }
    return RecipeRegistry.recipeFor(lookup).map(RecipeCatalogs::catalogEntryFor);
  }

  /** Returns one view-specific recipe detail payload by stable id, or empty when unknown. */
  public static Optional<RecipeCatalogDetail> lookupFor(String id) {
    java.util.Objects.requireNonNull(id, "id must not be null");
    String lookup = id.trim();
    if (lookup.isBlank()) {
      throw new IllegalArgumentException("id must not be blank");
    }
    return RecipeRegistry.recipeFor(lookup).map(RecipeCatalogs::catalogDetailFor);
  }

  private static List<RecipeCatalogEntry> recipeEntries() {
    return RecipeRegistry.recipes().stream().map(RecipeCatalogs::catalogEntryFor).toList();
  }

  private static RecipeCatalogEntry catalogEntryFor(Recipe recipe) {
    return new RecipeCatalogEntry(
        recipe.type(),
        recipe.id(),
        recipe.requestFileName(),
        recipe.summary(),
        recipe.advisory(),
        recipe.requiredWorkspacePaths());
  }

  static RecipeCatalogDetail catalogDetailFor(Recipe recipe) {
    RecipeRequestProfile requestProfile = requestProfileFor(recipe.plan());
    return switch (recipe.type()) {
      case EXAMPLE ->
          new ExampleRecipeCatalogDetail(
              recipe.id(),
              recipe.requestFileName(),
              recipe.summary(),
              recipe.advisory(),
              recipe.requiredWorkspacePaths(),
              recipe.intentTags(),
              requestProfile);
      case TASK_STARTER -> taskDetailFor(recipe, requestProfile);
    };
  }

  private static TaskStarterRecipeCatalogDetail taskDetailFor(
      Recipe recipe, RecipeRequestProfile requestProfile) {
    TaskEntry task = RecipeRegistry.taskEntryFor(recipe.id()).orElseThrow();
    return new TaskStarterRecipeCatalogDetail(
        task.id(),
        task.starter().requestFileName(),
        task.narrative().summary(),
        task.starter().advisory(),
        task.starter().requiredWorkspacePaths(),
        task.intentTags(),
        task.discoveryProfile(),
        task.narrative().outcomes(),
        task.narrative().requiredInputs(),
        task.narrative().optionalFeatures(),
        task.executionProfile(),
        task.interactionProfile(),
        task.workflow(),
        requestProfile);
  }

  private static RecipeRequestProfile requestProfileFor(WorkbookPlan plan) {
    Set<String> mutationActionTypes = new LinkedHashSet<>();
    Set<String> assertionTypes = new LinkedHashSet<>();
    Set<String> inspectionQueryTypes = new LinkedHashSet<>();
    for (var step : plan.steps()) {
      switch (step) {
        case MutationStep mutationStep ->
            mutationActionTypes.add(
                GridGrindProtocolTypeNames.mutationActionTypeName(
                    mutationStep.action().getClass()));
        case AssertionStep assertionStep ->
            assertionTypes.add(
                GridGrindProtocolTypeNames.assertionTypeName(assertionStep.assertion().getClass()));
        case InspectionStep inspectionStep ->
            inspectionQueryTypes.add(
                GridGrindProtocolTypeNames.inspectionQueryTypeName(
                    inspectionStep.query().getClass()));
      }
    }
    return new RecipeRequestProfile(
        GridGrindProtocolTypeNames.workbookSourceTypeName(plan.source().getClass()),
        GridGrindProtocolTypeNames.workbookPersistenceTypeName(plan.persistence().getClass()),
        plan.effectiveExecutionMode().modeType(),
        plan.journalLevel().name(),
        GridGrindProtocolTypeNames.calculationStrategyTypeName(
            plan.calculationPolicy().effectiveStrategy().getClass()),
        plan.calculationPolicy().markRecalculateOnOpen(),
        plan.steps().size(),
        List.copyOf(mutationActionTypes),
        List.copyOf(assertionTypes),
        List.copyOf(inspectionQueryTypes));
  }
}
