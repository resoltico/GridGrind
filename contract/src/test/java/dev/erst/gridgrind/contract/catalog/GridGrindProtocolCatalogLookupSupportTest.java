package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Focused tests for extracted protocol-catalog lookup and search helpers. */
class GridGrindProtocolCatalogLookupSupportTest {
  @Test
  void searchReturnsExactQualifiedMatchesBeforeFuzzyResults() {
    CatalogSearchResult result =
        GridGrindProtocolCatalogLookupSupport.search(
            lookupFixtureCatalog(), "  sourceTypes:SOURCE_ALPHA  ");

    assertEquals(result.matches().size(), result.totalCount());
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

    assertEquals(List.of("sourceTypes", "sourceTypes:SOURCE_ALPHA"), qualifiedIds(result));
  }

  @Test
  void searchCanMatchAcrossLookupIdAndSummaryText() {
    CatalogSearchResult result =
        GridGrindProtocolCatalogLookupSupport.search(lookupFixtureCatalog(), "source forecast");

    assertEquals(List.of("sourceTypes", "sourceTypes:SOURCE_ALPHA"), qualifiedIds(result));
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
  void searchMatchesSupportingEntriesBySummaryAndCollapsesThemBehindGroupResults() {
    Catalog catalog = lookupFixtureCatalog();

    CatalogSearchResult nestedSummary =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "legend variant");
    CatalogSearchResult plainSummary =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "display card");

    assertTrue(
        nestedSummary.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .anyMatch("nestedTypes:legendVariants"::equals));
    assertTrue(
        nestedSummary.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .noneMatch("legendVariants:LEGEND_ALPHA"::equals));
    assertTrue(
        plainSummary.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .anyMatch("plainTypes:displayCard"::equals));
    assertTrue(
        plainSummary.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .noneMatch("displayCard:DISPLAY_CARD"::equals));
  }

  @Test
  void searchMatchesSupportingEntriesThroughFieldTextBeforeCollapsingThem() {
    CatalogSearchResult result =
        GridGrindProtocolCatalogLookupSupport.search(lookupFixtureCatalog(), "title");

    assertTrue(
        result.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .anyMatch("nestedTypes:legendVariants"::equals));
    assertTrue(
        result.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .noneMatch("legendVariants:LEGEND_ALPHA"::equals));
  }

  @Test
  void searchPrioritizesOperationEntriesForConceptQueriesAndPublishesRelatedEntryIds() {
    CatalogSearchResult result = GridGrindProtocolCatalog.searchCatalog("chart title");

    assertTrue(
        result.matches().stream()
            .limit(2)
            .map(CatalogSearchMatch::qualifiedId)
            .anyMatch("mutationActionTypes:SET_CHART"::equals));
    assertTrue(
        result.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .noneMatch("chartTitleInputTypes:FORMULA"::equals));
    assertTrue(
        result.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .noneMatch("chartInputType:ChartInput"::equals));
    assertTrue(
        result.matches().stream()
            .filter(match -> "mutationActionTypes:SET_CHART".equals(match.qualifiedId()))
            .findFirst()
            .orElseThrow()
            .supportingMatches()
            .stream()
            .map(CatalogSearchMatch::qualifiedId)
            .anyMatch("nestedTypes:chartTitleInputTypes"::equals));
    assertTrue(
        result.matches().stream()
            .filter(match -> "mutationActionTypes:SET_CHART".equals(match.qualifiedId()))
            .findFirst()
            .orElseThrow()
            .supportingMatches()
            .stream()
            .map(CatalogSearchMatch::qualifiedId)
            .anyMatch("plainTypes:chartInputType"::equals));
  }

  @Test
  void searchResultRejectsNullMatches() {
    IllegalArgumentException exception =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new CatalogSearchResult(
                    GridGrindProtocolVersion.current(),
                    "source",
                    new java.util.ArrayList<>(java.util.Arrays.asList((CatalogSearchMatch) null))));

    assertEquals("matches must not contain nulls", exception.getMessage());
  }

  @Test
  void lookupValueHelpersReturnTopLevelTypeGroupForCategoryNames() {
    Object mutationGroup =
        GridGrindProtocolCatalogLookupSupport.lookupValueFor(
                GridGrindProtocolCatalog.catalog(), "mutationActionTypes")
            .orElseThrow();
    Object sourceGroup =
        GridGrindProtocolCatalogLookupSupport.lookupValueFor(lookupFixtureCatalog(), "sourceTypes")
            .orElseThrow();

    assertTrue(mutationGroup instanceof TopLevelTypeGroup);
    assertTrue(sourceGroup instanceof TopLevelTypeGroup);
    assertFalse(((TopLevelTypeGroup) mutationGroup).types().isEmpty());
    assertEquals(1, ((TopLevelTypeGroup) sourceGroup).types().size());
  }

  @Test
  void lookupValueHelpersReturnPublishedEntryAndGroupShapes() {
    Object nestedGroup =
        GridGrindProtocolCatalogLookupSupport.lookupValueFor(
                lookupFixtureCatalog(), "nestedTypes:legendVariants")
            .orElseThrow();
    Object plainGroup =
        GridGrindProtocolCatalogLookupSupport.lookupValueFor(
                lookupFixtureCatalog(), "plainTypes:displayCard")
            .orElseThrow();
    Object sourceEntry =
        GridGrindProtocolCatalogLookupSupport.lookupValueFor(
                lookupFixtureCatalog(), "sourceTypes:SOURCE_ALPHA")
            .orElseThrow();

    assertTrue(nestedGroup instanceof NestedTypeGroup);
    assertTrue(plainGroup instanceof PlainTypeGroup);
    assertTrue(sourceEntry instanceof TypeEntry);
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
        List.of(new TypeEntry("SOURCE_ALPHA", "Budget forecast entry.", List.of())),
        List.of(),
        List.of(),
        List.of(),
        List.of(),
        List.of(),
        List.of(new NestedTypeGroup("legendVariants", "type", List.of(nestedLegend()))),
        List.of(new PlainTypeGroup("displayCard", displayCard())),
        List.of());
  }

  private static TypeEntry nestedLegend() {
    return new TypeEntry(
        "LEGEND_ALPHA",
        "Legend variant entry.",
        List.of(
            new FieldEntry(
                "title",
                FieldRequirement.REQUIRED,
                new FieldShape.Scalar(ScalarType.STRING),
                List.of())));
  }

  private static TypeEntry displayCard() {
    return new TypeEntry(
        "DISPLAY_CARD",
        "Display card summary.",
        List.of(
            new FieldEntry(
                "label",
                FieldRequirement.REQUIRED,
                new FieldShape.Scalar(ScalarType.STRING),
                List.of())));
  }
}
