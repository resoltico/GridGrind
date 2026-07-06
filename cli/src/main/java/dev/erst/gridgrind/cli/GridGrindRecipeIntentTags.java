package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.examples.GridGrindCliRecipeRegistry;
import java.util.List;

/** Shared published recipe intent-tag vocabulary used by discovery surfaces. */
final class GridGrindRecipeIntentTags {
  private static final List<String> PUBLISHED_INTENT_TAGS =
      GridGrindCliRecipeRegistry.recipes().stream()
          .flatMap(recipe -> recipe.intentTags().stream())
          .distinct()
          .sorted()
          .toList();

  private GridGrindRecipeIntentTags() {}

  static List<String> publishedIntentTags() {
    return PUBLISHED_INTENT_TAGS;
  }
}
