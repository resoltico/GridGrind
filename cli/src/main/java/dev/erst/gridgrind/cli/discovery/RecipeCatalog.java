package dev.erst.gridgrind.cli.discovery;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.List;

/** Unified machine-readable catalog of every built-in recipe the CLI can print. */
public record RecipeCatalog(
    GridGrindProtocolVersion protocolVersion, List<RecipeCatalogEntry> recipes) {
  public RecipeCatalog {
    protocolVersion = CliDiscoveryValidation.requireProtocolVersion(protocolVersion);
    recipes = CliRecipeCatalogValidation.copyRecipeCatalogEntries(recipes, "recipes");
  }
}
