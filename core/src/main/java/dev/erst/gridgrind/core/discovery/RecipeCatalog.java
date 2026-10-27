package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.util.List;

/** Unified machine-readable catalog of every published built-in recipe. */
public record RecipeCatalog(ProtocolVersion protocolVersion, List<RecipeCatalogEntry> recipes) {
  public RecipeCatalog {
    protocolVersion = DiscoveryValidation.requireProtocolVersion(protocolVersion);
    recipes = RecipeCatalogValidation.copyRecipeCatalogEntries(recipes, "recipes");
  }
}
