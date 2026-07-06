package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.cli.discovery.ExampleWorkspaceMode;
import dev.erst.gridgrind.cli.discovery.RecipeView;
import dev.erst.gridgrind.cli.discovery.ShippedExampleCatalog;
import dev.erst.gridgrind.cli.discovery.ShippedExampleEntry;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

/** Built-in example views derived from the canonical CLI recipe registry. */
public final class GridGrindShippedExamples {
  /** One built-in example request emitted by the CLI and mirrored under `examples/`. */
  public record ShippedExample(
      String id, String requestFileName, String summary, WorkbookPlan plan) {
    public ShippedExample {
      id = requireNonBlank(id, "id");
      if (!id.equals(id.toUpperCase(Locale.ROOT))) {
        throw new IllegalArgumentException("id must use upper-case discovery tokens");
      }
      requestFileName = requireNonBlank(requestFileName, "requestFileName");
      summary = requireNonBlank(summary, "summary");
      Objects.requireNonNull(plan, "plan must not be null");
      if (!requestFileName.endsWith(".json")) {
        throw new IllegalArgumentException("requestFileName must end with .json");
      }
    }
  }

  /** Indicates whether one built-in example is portable or requires repository assets. */
  public record ExampleRequirements(
      ExampleWorkspaceMode workspaceMode, List<String> requiredWorkspacePaths) {
    public ExampleRequirements {
      Objects.requireNonNull(workspaceMode, "workspaceMode must not be null");
      requiredWorkspacePaths =
          List.copyOf(
              Objects.requireNonNull(
                  requiredWorkspacePaths, "requiredWorkspacePaths must not be null"));
    }
  }

  private GridGrindShippedExamples() {}

  /** Returns the ordered list of built-in examples. */
  public static List<ShippedExample> examples() {
    return exampleRecipes().stream().map(GridGrindShippedExamples::toShippedExample).toList();
  }

  /**
   * Returns the checked-in example fixtures rooted for in-repository execution from `examples/`.
   */
  public static List<ShippedExample> repositoryExamples() {
    return GridGrindCliRecipeRegistry.repositoryExamples();
  }

  /** Returns built-in examples that can execute from a blank artifact workspace. */
  public static List<ShippedExample> selfContainedExamples() {
    return examples().stream()
        .filter(
            example ->
                requirementsFor(example).workspaceMode() == ExampleWorkspaceMode.SELF_CONTAINED)
        .toList();
  }

  /** Returns built-in examples that require copied repository asset directories. */
  public static List<ShippedExample> repositoryAssetBackedExamples() {
    return examples().stream()
        .filter(
            example ->
                requirementsFor(example).workspaceMode()
                    == ExampleWorkspaceMode.REQUIRES_EXAMPLE_ASSETS)
        .toList();
  }

  /** Returns public catalog metadata for the built-in example set. */
  public static List<ShippedExampleEntry> catalogEntries() {
    return GridGrindCliRecipeRegistry.exampleCatalogEntries();
  }

  /** Returns the machine-readable example catalog for CLI discovery. */
  public static ShippedExampleCatalog catalog() {
    return new ShippedExampleCatalog(
        dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion.current(), catalogEntries());
  }

  /** Finds one built-in example by its stable upper-case id. */
  public static Optional<ShippedExample> find(String id) {
    Objects.requireNonNull(id, "id must not be null");
    return GridGrindCliRecipeRegistry.recipeFor(id)
        .filter(recipe -> recipe.view() == RecipeView.EXAMPLE)
        .map(GridGrindShippedExamples::toShippedExample);
  }

  /** Returns the portability contract for one stable built-in example id. */
  public static Optional<ExampleWorkspaceMode> workspaceModeFor(String id) {
    Objects.requireNonNull(id, "id must not be null");
    return GridGrindCliRecipeRegistry.recipeFor(id)
        .filter(recipe -> recipe.view() == RecipeView.EXAMPLE)
        .map(GridGrindCliRecipe::workspaceMode);
  }

  /** Returns the portability requirements for one concrete built-in example entry. */
  public static ExampleRequirements requirementsFor(ShippedExample example) {
    Objects.requireNonNull(example, "example must not be null");
    GridGrindCliRecipe recipe =
        GridGrindCliRecipeRegistry.recipeFor(example.id())
            .filter(candidate -> candidate.view() == RecipeView.EXAMPLE)
            .orElseThrow(
                () ->
                    new IllegalStateException(
                        "Missing shipped-example requirements for " + example.id()));
    return new ExampleRequirements(recipe.workspaceMode(), recipe.requiredWorkspacePaths());
  }

  private static List<GridGrindCliRecipe> exampleRecipes() {
    return GridGrindCliRecipeRegistry.exampleRecipes();
  }

  private static ShippedExample toShippedExample(GridGrindCliRecipe recipe) {
    return new ShippedExample(
        recipe.id(), recipe.requestFileName(), recipe.summary(), recipe.plan());
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
