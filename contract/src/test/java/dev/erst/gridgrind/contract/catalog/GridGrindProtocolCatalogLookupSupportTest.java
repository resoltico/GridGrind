package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.List;
import org.junit.jupiter.api.Test;

/** Focused tests for extracted protocol-catalog lookup and search helpers. */
class GridGrindProtocolCatalogLookupSupportTest {
  @Test
  void searchReturnsExactQualifiedMatchesBeforeFuzzyResults() {
    CatalogSearchResult result =
        GridGrindProtocolCatalogLookupSupport.search(
            lookupFixtureCatalog(), "  sourceTypes:SOURCE_ALPHA  ");

    assertEquals("sourceTypes:SOURCE_ALPHA", result.matches().getFirst().qualifiedId());
    assertEquals("ENTRY", result.matches().getFirst().kind());
  }

  @Test
  void searchReturnsExactLookupIdMatchesBeforeQualifiedFallbacks() {
    CatalogSearchResult result =
        GridGrindProtocolCatalogLookupSupport.search(lookupFixtureCatalog(), "source_alpha");

    assertEquals("sourceTypes:SOURCE_ALPHA", result.matches().getFirst().qualifiedId());
    assertEquals("ENTRY", result.matches().getFirst().kind());
  }

  @Test
  void searchMatchesQualifiedIdsWhenLookupIdsDoNotContainTheQuery() {
    CatalogSearchResult result =
        GridGrindProtocolCatalogLookupSupport.search(lookupFixtureCatalog(), "sourcetypes");

    assertEquals(List.of("sourceTypes:SOURCE_ALPHA"), qualifiedIds(result));
  }

  @Test
  void searchCanMatchAcrossLookupIdAndSummaryText() {
    CatalogSearchResult result =
        GridGrindProtocolCatalogLookupSupport.search(lookupFixtureCatalog(), "source forecast");

    assertEquals(List.of("sourceTypes:SOURCE_ALPHA"), qualifiedIds(result));
  }

  @Test
  void searchReturnsTypedResultsForNestedAndPlainGroups() {
    CatalogSearchResult nestedResult =
        GridGrindProtocolCatalogLookupSupport.search(
            lookupFixtureCatalog(), "nestedTypes:legendVariants");
    CatalogSearchResult plainResult =
        GridGrindProtocolCatalogLookupSupport.search(
            lookupFixtureCatalog(), "plainTypes:displayCard");

    assertEquals("nestedTypes:legendVariants", nestedResult.matches().getFirst().qualifiedId());
    assertEquals("NESTED_GROUP", nestedResult.matches().getFirst().kind());
    assertEquals("plainTypes:displayCard", plainResult.matches().getFirst().qualifiedId());
    assertEquals("PLAIN_GROUP", plainResult.matches().getFirst().kind());
  }

  @Test
  void searchResultRejectsNullMatches() {
    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new CatalogSearchResult(
                    "source",
                    new java.util.ArrayList<>(java.util.Arrays.asList((CatalogSearchMatch) null))));

    assertEquals("matches must not contain nulls", exception.getMessage());
  }

  @Test
  void lookupHelpersRejectUnsupportedValueKinds() {
    IllegalArgumentException kindFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> GridGrindProtocolCatalogLookupSupport.kindFor("bogus-kind"));
    assertEquals("Unsupported catalog lookup value: bogus-kind", kindFailure.getMessage());

    IllegalArgumentException summaryFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> GridGrindProtocolCatalogLookupSupport.summaryFor("bogus-summary"));
    assertEquals("Unsupported catalog lookup value: bogus-summary", summaryFailure.getMessage());
  }

  private static List<String> qualifiedIds(CatalogSearchResult result) {
    return result.matches().stream().map(CatalogSearchMatch::qualifiedId).toList();
  }

  private static Catalog lookupFixtureCatalog() {
    Catalog published = GridGrindProtocolCatalog.catalog();
    return new Catalog(
        published.protocolVersion(),
        published.discriminatorField(),
        new TypeEntry("REQUEST", "Synthetic request catalog entry.", List.of()),
        published.cliSurface(),
        List.of(),
        List.of(new TypeEntry("SOURCE_ALPHA", "Budget forecast entry.", List.of())),
        List.of(),
        List.of(),
        List.of(),
        List.of(),
        List.of(),
        List.of(new NestedTypeGroup("legendVariants", "type", List.of(nestedLegend()))),
        List.of(new PlainTypeGroup("displayCard", displayCard())));
  }

  private static TypeEntry nestedLegend() {
    return new TypeEntry("LEGEND_ALPHA", "Legend variant entry.", List.of());
  }

  private static TypeEntry displayCard() {
    return new TypeEntry("DISPLAY_CARD", "Display card summary.", List.of());
  }
}
