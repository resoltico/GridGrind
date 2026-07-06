package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.cli.examples.GridGrindCliRecipe;
import dev.erst.gridgrind.cli.examples.GridGrindCliRecipeRegistry;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.List;
import java.util.Optional;

/** Unified public recipe catalog layered on top of the canonical recipe registry. */
public final class GridGrindRecipeCatalog {
  private GridGrindRecipeCatalog() {}

  /** Returns the full built-in recipe catalog across examples and task starters. */
  public static RecipeCatalog catalog() {
    return new RecipeCatalog(GridGrindProtocolVersion.current(), recipeEntries());
  }

  /** Returns one recipe catalog entry by stable id, or empty when unknown. */
  public static Optional<RecipeCatalogEntry> entryFor(String id) {
    java.util.Objects.requireNonNull(id, "id must not be null");
    String lookup = id.trim();
    if (lookup.isBlank()) {
      throw new IllegalArgumentException("id must not be blank");
    }
    return GridGrindCliRecipeRegistry.recipeFor(lookup)
        .map(GridGrindRecipeCatalog::catalogEntryFor);
  }

  /** Returns one view-specific recipe detail payload by stable id, or empty when unknown. */
  public static Optional<RecipeCatalogDetail> lookupFor(String id) {
    java.util.Objects.requireNonNull(id, "id must not be null");
    String lookup = id.trim();
    if (lookup.isBlank()) {
      throw new IllegalArgumentException("id must not be blank");
    }
    return GridGrindCliRecipeRegistry.recipeFor(lookup)
        .map(RecipeCatalogDetailSupport::catalogDetailFor);
  }

  private static List<RecipeCatalogEntry> recipeEntries() {
    return GridGrindCliRecipeRegistry.recipes().stream()
        .map(GridGrindRecipeCatalog::catalogEntryFor)
        .toList();
  }

  private static RecipeCatalogEntry catalogEntryFor(GridGrindCliRecipe recipe) {
    return new RecipeCatalogEntry(
        recipe.view(),
        recipe.id(),
        recipe.requestFileName(),
        recipe.summary(),
        recipe.workspaceMode(),
        recipe.requiredWorkspacePaths());
  }
}
