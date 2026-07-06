package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.cli.discovery.RecipeView;
import dev.erst.gridgrind.cli.discovery.ShippedExampleEntry;
import dev.erst.gridgrind.cli.discovery.TaskEntry;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/** Single authoritative recipe registry backing CLI example and task-starter views. */
public final class GridGrindCliRecipeRegistry {
  private static final RecipeRegistryState REGISTRY = buildRegistryState();
  private static final List<GridGrindCliRecipe> RECIPES = List.copyOf(REGISTRY.recipes());
  private static final List<GridGrindCliRecipe> EXAMPLE_RECIPES =
      RECIPES.stream().filter(recipe -> recipe.view() == RecipeView.EXAMPLE).toList();
  private static final List<GridGrindCliRecipe> TASK_RECIPES =
      RECIPES.stream().filter(recipe -> recipe.view() == RecipeView.TASK_STARTER).toList();
  private static final Map<String, GridGrindCliRecipe> RECIPES_BY_ID =
      Map.copyOf(REGISTRY.recipesById());
  private static final Map<String, WorkbookPlan> REPOSITORY_EXAMPLE_PLANS =
      Map.copyOf(REGISTRY.repositoryExamplePlans());
  private static final Map<String, TaskEntry> TASK_ENTRIES = Map.copyOf(REGISTRY.taskEntries());

  private GridGrindCliRecipeRegistry() {}

  /** Returns the ordered example view over the canonical recipe source. */
  static List<GridGrindCliRecipe> exampleRecipes() {
    return EXAMPLE_RECIPES;
  }

  /** Returns the ordered task-starter view over the canonical recipe source. */
  static List<GridGrindCliRecipe> taskRecipes() {
    return TASK_RECIPES;
  }

  /** Returns the full ordered recipe registry across examples and task starters. */
  public static List<GridGrindCliRecipe> recipes() {
    return RECIPES;
  }

  /** Resolves one published recipe by its stable upper-case id. */
  public static Optional<GridGrindCliRecipe> recipeFor(String id) {
    Objects.requireNonNull(id, "id must not be null");
    return Optional.ofNullable(RECIPES_BY_ID.get(id));
  }

  static List<ShippedExampleEntry> exampleCatalogEntries() {
    return EXAMPLE_RECIPES.stream()
        .map(
            recipe ->
                new ShippedExampleEntry(
                    recipe.id(),
                    recipe.requestFileName(),
                    recipe.summary(),
                    recipe.workspaceMode(),
                    recipe.requiredWorkspacePaths()))
        .toList();
  }

  static List<GridGrindShippedExamples.ShippedExample> repositoryExamples() {
    return EXAMPLE_RECIPES.stream()
        .map(
            recipe ->
                new GridGrindShippedExamples.ShippedExample(
                    recipe.id(),
                    recipe.requestFileName(),
                    recipe.summary(),
                    repositoryExamplePlanFor(recipe.id())))
        .toList();
  }

  /** Returns the ordered task-entry view over the canonical recipe source. */
  public static List<TaskEntry> taskEntries() {
    return TASK_RECIPES.stream().map(recipe -> requireTaskEntry(recipe.id())).toList();
  }

  /** Resolves one published task entry by its stable upper-case id. */
  public static Optional<TaskEntry> taskEntryFor(String id) {
    Objects.requireNonNull(id, "id must not be null");
    return Optional.ofNullable(TASK_ENTRIES.get(id));
  }

  private static RecipeRegistryState buildRegistryState() {
    RecipeRegistryState state = new RecipeRegistryState();
    registerDefinitions(state);
    return state;
  }

  private static void registerDefinitions(RecipeRegistryState state) {
    for (GridGrindRecipeDefinition definition : GridGrindRecipeDefinitions.definitions()) {
      registerDefinition(state, definition);
    }
  }

  private static void registerDefinition(
      RecipeRegistryState state, GridGrindRecipeDefinition definition) {
    if (definition instanceof GridGrindExampleRecipeDefinition example) {
      registerRecipe(state, example.recipe());
      putUniqueAssociated(
          state.repositoryExamplePlans(),
          example.id(),
          example.repositoryPlan(),
          "repository example plan");
      return;
    }
    GridGrindTaskRecipeDefinition task = (GridGrindTaskRecipeDefinition) definition;
    registerRecipe(state, task.recipe());
    putUniqueAssociated(state.taskEntries(), task.task().id(), task.task(), "task entry");
  }

  private static void registerRecipe(RecipeRegistryState state, GridGrindCliRecipe recipe) {
    state.recipes().add(recipe);
    putUniqueRecipe(state.recipesById(), recipe);
  }

  static WorkbookPlan repositoryExamplePlanFor(String id) {
    WorkbookPlan repositoryPlan = REPOSITORY_EXAMPLE_PLANS.get(id);
    if (repositoryPlan == null) {
      throw new IllegalStateException("Missing repository example plan for " + id);
    }
    return repositoryPlan;
  }

  static TaskEntry requireTaskEntry(String id) {
    TaskEntry taskEntry = TASK_ENTRIES.get(id);
    if (taskEntry == null) {
      throw new IllegalStateException("Missing task entry for recipe " + id);
    }
    return taskEntry;
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  static Map<String, GridGrindCliRecipe> recipesById(List<GridGrindCliRecipe> recipes) {
    Map<String, GridGrindCliRecipe> byId = new LinkedHashMap<>();
    for (GridGrindCliRecipe recipe : recipes) {
      putUniqueRecipe(byId, recipe);
    }
    return Map.copyOf(byId);
  }

  private static void putUniqueRecipe(
      Map<String, GridGrindCliRecipe> byId, GridGrindCliRecipe recipe) {
    GridGrindCliRecipe previous = byId.put(recipe.id(), recipe);
    if (previous != null) {
      throw new IllegalStateException("Duplicate CLI recipe id " + recipe.id());
    }
  }

  static <T> void putUniqueAssociated(Map<String, T> byId, String id, T value, String label) {
    T previous = byId.put(id, value);
    if (previous != null) {
      throw new IllegalStateException("Duplicate " + label + " for " + id);
    }
  }

  private record RecipeRegistryState(
      List<GridGrindCliRecipe> recipes,
      Map<String, GridGrindCliRecipe> recipesById,
      Map<String, WorkbookPlan> repositoryExamplePlans,
      Map<String, TaskEntry> taskEntries) {
    private RecipeRegistryState() {
      this(new ArrayList<>(), new LinkedHashMap<>(), new LinkedHashMap<>(), new LinkedHashMap<>());
    }
  }
}
