package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.discovery.ProtocolCatalogFieldMetadataKey;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogGroupIndex;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogIndexReport;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogLookupNamespace;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogSearchHit;
import dev.erst.gridgrind.cli.discovery.ProtocolCatalogSearchReport;
import dev.erst.gridgrind.cli.examples.GridGrindCliRecipe;
import dev.erst.gridgrind.cli.examples.GridGrindCliRecipeRegistry;
import dev.erst.gridgrind.contract.catalog.Catalog;
import dev.erst.gridgrind.contract.catalog.CatalogSearchMatch;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolCatalog;
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
          + "? Run gridgrind --print-recipe-catalog to list recipe ids or"
          + " gridgrind --print-recipe-keyword-match --query \"monthly sales dashboard\""
          + " when you know the goal but not the id.";
    }
    return "Unknown recipe: "
        + recipeId
        + ". Run gridgrind --print-recipe-catalog or"
        + " gridgrind --print-recipe-keyword-match --query \"monthly sales dashboard\""
        + " to discover valid recipe ids.";
  }

  static String unknownOperationMessage(String operationId) {
    return CliSuggestionSupport.protocolCatalogSearchCommandForLookupId(operationId)
        .map(
            command ->
                "Unknown lookup id: "
                    + operationId
                    + ". Run "
                    + command
                    + " or gridgrind --print-protocol-catalog to discover valid lookup ids.")
        .orElse(
            "Unknown lookup id: "
                + operationId
                + ". Run gridgrind --print-protocol-catalog to discover valid lookup ids.");
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
        List.of(
            new ProtocolCatalogFieldMetadataKey(
                "projectedByFacets",
                "Field is present only when one of the listed read-projection facets is"
                    + " requested."),
            new ProtocolCatalogFieldMetadataKey(
                "noteRefs",
                "Entry references stable note ids published once in the scoped lookup payload or"
                    + " full catalog."),
            new ProtocolCatalogFieldMetadataKey(
                "enumValueDocs",
                "Field publishes per-enum-value summaries aligned with enumValues so agents can"
                    + " choose the right token directly.")),
        List.of(
            new ProtocolCatalogLookupNamespace(
                "<topLevelGroup>:<id>",
                "Resolve one top-level protocol type such as mutationActionTypes:SET_CELL."),
            new ProtocolCatalogLookupNamespace(
                "nestedTypes:<group>",
                "Resolve one nested tagged-union group such as nestedTypes:cellInputTypes."),
            new ProtocolCatalogLookupNamespace(
                "plainTypes:<group>",
                "Resolve one plain record group such as plainTypes:sheetSummaryReport."),
            new ProtocolCatalogLookupNamespace(
                "<id>",
                "Resolve one unqualified top-level type id only when it is globally unique.")));
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

  private static List<String> typeIds(List<dev.erst.gridgrind.contract.catalog.TypeEntry> types) {
    return types.stream().map(dev.erst.gridgrind.contract.catalog.TypeEntry::id).toList();
  }

  private static Optional<String> suggestedRecipeId(String recipeId) {
    String normalizedRecipeId = normalizeLookupToken(recipeId);
    return GridGrindCliRecipeRegistry.recipes().stream()
        .filter(recipe -> recipeLooksLikeLookup(recipe, recipeId, normalizedRecipeId))
        .map(GridGrindCliRecipe::id)
        .findFirst();
  }

  private static boolean recipeLooksLikeLookup(
      GridGrindCliRecipe recipe, String lookupText, String normalizedLookupText) {
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
