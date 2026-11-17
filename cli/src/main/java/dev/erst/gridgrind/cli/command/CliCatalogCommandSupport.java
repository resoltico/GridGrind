package dev.erst.gridgrind.cli.command;

import dev.erst.gridgrind.cli.render.ProtocolCatalogGroupIndex;
import dev.erst.gridgrind.cli.render.ProtocolCatalogIndexReport;
import dev.erst.gridgrind.cli.render.ProtocolCatalogSearchHit;
import dev.erst.gridgrind.cli.render.ProtocolCatalogSearchReport;
import dev.erst.gridgrind.core.catalog.Catalog;
import dev.erst.gridgrind.core.catalog.CatalogSearchMatch;
import dev.erst.gridgrind.core.catalog.GridGrindProtocolCatalog;
import dev.erst.gridgrind.core.discovery.Recipe;
import dev.erst.gridgrind.core.discovery.RecipeRegistry;
import java.util.List;
import java.util.Locale;
import java.util.Optional;

/** Shared discovery, warning, and search-summary helpers for CLI catalog commands. */
final class CliCatalogCommandSupport {
  private CliCatalogCommandSupport() {}

  static String unknownRecipeMessage(String recipeId) {
    Optional<String> recipeSuggestion = suggestedRecipeId(recipeId);
    if (recipeSuggestion.isPresent()) {
      return "Unknown recipe: "
          + recipeId
          + ". Recipe ids use stable upper-case tokens; did you mean "
          + recipeSuggestion.orElseThrow()
          + "? Run gridgrind recipe list to list recipe ids or"
          + " gridgrind recipe match \"monthly sales dashboard\""
          + " when you know the goal but not the id.";
    }
    return "Unknown recipe: "
        + recipeId
        + ". Run gridgrind recipe list or"
        + " gridgrind recipe match \"monthly sales dashboard\""
        + " to discover valid recipe ids.";
  }

  static String unknownOperationMessage(String operationId) {
    return CliSuggestionSupport.protocolCatalogSearchCommandForLookupId(operationId)
        .map(command ->
            "Unknown lookup id: "
                + operationId
                + ". Run "
                + command
                + " or gridgrind catalog to discover valid lookup ids.")
        .orElse(
            "Unknown lookup id: "
                + operationId
                + ". Run gridgrind catalog to discover valid lookup ids.");
  }

  static ProtocolCatalogIndexReport protocolCatalogIndexReport() {
    Catalog catalog = GridGrindProtocolCatalog.catalog();
    return new ProtocolCatalogIndexReport(
        catalog.protocolVersion(),
        catalog.discriminatorField(),
        catalog.requestType().id(),
        catalog.topLevelGroups().stream()
            .map(group -> new ProtocolCatalogGroupIndex(group.group(), typeIds(group.types())))
            .toList(),
        catalog.nestedTypes().stream()
            .map(group -> new ProtocolCatalogGroupIndex(group.group(), typeIds(group.types())))
            .toList(),
        catalog.plainTypes().stream()
            .map(group -> new ProtocolCatalogGroupIndex(group.group(), List.of(group.type().id())))
            .toList(),
        GridGrindProtocolCatalog.fieldMetadataKeys(),
        GridGrindProtocolCatalog.lookupNamespaces(),
        catalog.limits());
  }

  static ProtocolCatalogSearchReport summarizedSearchReport(String query) {
    var result = GridGrindProtocolCatalog.searchCatalog(query);
    return new ProtocolCatalogSearchReport(
        result.protocolVersion(),
        result.query(),
        result.matches().stream().map(CliCatalogCommandSupport::summarizedSearchHit).toList());
  }

  private static ProtocolCatalogSearchHit summarizedSearchHit(CatalogSearchMatch match) {
    return new ProtocolCatalogSearchHit(
        match.catalogGroup(),
        match.lookupId(),
        match.qualifiedId(),
        match.kind(),
        match.summary(),
        match.relatedEntryIds(),
        match.supportingMatches().stream().map(CatalogSearchMatch::qualifiedId).toList());
  }

  private static List<String> typeIds(List<dev.erst.gridgrind.core.catalog.TypeEntry> types) {
    return types.stream().map(dev.erst.gridgrind.core.catalog.TypeEntry::id).toList();
  }

  private static Optional<String> suggestedRecipeId(String recipeId) {
    String normalizedRecipeId = normalizeLookupToken(recipeId);
    return RecipeRegistry.recipes().stream()
        .filter(recipe -> recipeLooksLikeLookup(recipe, recipeId, normalizedRecipeId))
        .map(Recipe::id)
        .findFirst();
  }

  private static boolean recipeLooksLikeLookup(
      Recipe recipe, String lookupText, String normalizedLookupText) {
    if (recipe.id().equalsIgnoreCase(lookupText)
        || normalizeLookupToken(recipe.id()).equals(normalizedLookupText)) {
      return true;
    }
    return recipe.requestFileName().equalsIgnoreCase(lookupText)
        || exampleStem(recipe.requestFileName()).equalsIgnoreCase(lookupText)
        || normalizeLookupToken(exampleStem(recipe.requestFileName())).equals(normalizedLookupText);
  }

  private static String normalizeLookupToken(String value) {
    return value.toUpperCase(Locale.ROOT).replace('-', '_').replace(' ', '_');
  }

  private static String exampleStem(String fileName) {
    return fileName.substring(0, fileName.length() - 5);
  }
}
