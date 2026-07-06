package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.cli.examples.GridGrindCliRecipe;
import dev.erst.gridgrind.cli.examples.GridGrindCliRecipeRegistry;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

/** Builds view-specific recipe-catalog lookup payloads from the canonical recipe registry. */
final class RecipeCatalogDetailSupport {
  private RecipeCatalogDetailSupport() {}

  static RecipeCatalogDetail catalogDetailFor(GridGrindCliRecipe recipe) {
    RecipeRequestProfile requestProfile = requestProfileFor(recipe.plan());
    return switch (recipe.view()) {
      case EXAMPLE ->
          new ExampleRecipeCatalogDetail(
              recipe.id(),
              recipe.requestFileName(),
              recipe.summary(),
              recipe.workspaceMode(),
              recipe.requiredWorkspacePaths(),
              recipe.intentTags(),
              requestProfile);
      case TASK_STARTER -> taskDetailFor(recipe, requestProfile);
    };
  }

  private static TaskStarterRecipeCatalogDetail taskDetailFor(
      GridGrindCliRecipe recipe, RecipeRequestProfile requestProfile) {
    TaskEntry task = GridGrindCliRecipeRegistry.taskEntryFor(recipe.id()).orElseThrow();
    return new TaskStarterRecipeCatalogDetail(
        task.id(),
        task.starter().requestFileName(),
        task.narrative().summary(),
        task.starter().workspaceMode(),
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
