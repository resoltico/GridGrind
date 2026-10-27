package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

/** Built-in example views derived from the canonical recipe registry. */
public final class ShippedExamples {
  /** One built-in example request mirrored under `examples/`. */
  public record ShippedExample(
      String id, String requestFileName, String summary, WorkbookPlan plan) {
    public ShippedExample {
      id = TextValidation.requireNonBlank(id, "id");
      if (!id.equals(id.toUpperCase(Locale.ROOT))) {
        throw new IllegalArgumentException("id must use upper-case discovery tokens");
      }
      requestFileName = TextValidation.requireNonBlank(requestFileName, "requestFileName");
      summary = TextValidation.requireNonBlank(summary, "summary");
      Objects.requireNonNull(plan, "plan must not be null");
      if (!requestFileName.endsWith(".json")) {
        throw new IllegalArgumentException("requestFileName must end with .json");
      }
    }
  }

  /** Indicates whether one built-in example is portable or requires repository assets. */
  public record ExampleRequirements(RecipeAdvisory advisory, List<String> requiredWorkspacePaths) {
    public ExampleRequirements {
      Objects.requireNonNull(advisory, "advisory must not be null");
      requiredWorkspacePaths =
          List.copyOf(
              Objects.requireNonNull(
                  requiredWorkspacePaths, "requiredWorkspacePaths must not be null"));
    }
  }

  private ShippedExamples() {}

  /** Returns the ordered list of built-in examples. */
  public static List<ShippedExample> examples() {
    return exampleRecipes().stream().map(ShippedExamples::toShippedExample).toList();
  }

  /**
   * Returns the checked-in example fixtures rooted for in-repository execution from `examples/`.
   */
  public static List<ShippedExample> repositoryExamples() {
    return RecipeRegistry.repositoryExamples();
  }

  /** Returns built-in examples that can execute from a blank artifact workspace. */
  public static List<ShippedExample> selfContainedExamples() {
    return examples().stream()
        .filter(example -> requirementsFor(example).advisory() == RecipeAdvisory.SELF_CONTAINED)
        .toList();
  }

  /** Returns built-in examples that require copied repository asset directories. */
  public static List<ShippedExample> repositoryAssetBackedExamples() {
    return examples().stream()
        .filter(
            example ->
                requirementsFor(example).advisory() == RecipeAdvisory.REQUIRES_EXAMPLE_ASSETS)
        .toList();
  }

  /** Returns public catalog metadata for the built-in example set. */
  public static List<ShippedExampleEntry> catalogEntries() {
    return RecipeRegistry.exampleCatalogEntries();
  }

  /** Returns the machine-readable example catalog for product discovery. */
  public static ShippedExampleCatalog catalog() {
    return new ShippedExampleCatalog(
        dev.erst.gridgrind.core.vocabulary.ProtocolVersion.CURRENT, catalogEntries());
  }

  /** Finds one built-in example by its stable upper-case id. */
  public static Optional<ShippedExample> find(String id) {
    Objects.requireNonNull(id, "id must not be null");
    return RecipeRegistry.recipeFor(id)
        .filter(recipe -> recipe.type() == RecipeType.EXAMPLE)
        .map(ShippedExamples::toShippedExample);
  }

  /** Returns the portability contract for one stable built-in example id. */
  public static Optional<RecipeAdvisory> advisoryFor(String id) {
    Objects.requireNonNull(id, "id must not be null");
    return RecipeRegistry.recipeFor(id)
        .filter(recipe -> recipe.type() == RecipeType.EXAMPLE)
        .map(Recipe::advisory);
  }

  /** Returns the portability requirements for one concrete built-in example entry. */
  public static ExampleRequirements requirementsFor(ShippedExample example) {
    Objects.requireNonNull(example, "example must not be null");
    Recipe recipe =
        RecipeRegistry.recipeFor(example.id())
            .filter(candidate -> candidate.type() == RecipeType.EXAMPLE)
            .orElseThrow(
                () ->
                    new IllegalStateException(
                        "Missing shipped-example requirements for " + example.id()));
    return new ExampleRequirements(recipe.advisory(), recipe.requiredWorkspacePaths());
  }

  private static List<Recipe> exampleRecipes() {
    return RecipeRegistry.exampleRecipes();
  }

  private static ShippedExample toShippedExample(Recipe recipe) {
    return new ShippedExample(
        recipe.id(), recipe.requestFileName(), recipe.summary(), recipe.plan());
  }
}
