package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import java.util.List;

/** Unified machine-readable catalog of every built-in recipe the CLI can print. */
public record RecipeCatalog(ProtocolVersion protocolVersion, List<RecipeCatalogEntry> recipes) {
  public RecipeCatalog {
    protocolVersion = CliDiscoveryValidation.requireProtocolVersion(protocolVersion);
    recipes = CliRecipeCatalogValidation.copyRecipeCatalogEntries(recipes, "recipes");
  }
}
