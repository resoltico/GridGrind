package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/**
 * Coverage tests for grouped lookup/search behavior introduced in the extracted catalog helpers.
 */
class CatalogLookupSupportCoverageTest {
  @Test
  void searchRelatesNestedAndPlainGroupsBackToTopLevelEntriesAndRejectsUnknownTypeSets() {
    Catalog catalog = lookupCoverageCatalog();

    CatalogSearchResult nested =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "nestedTypes:lookupNested");
    CatalogSearchResult plain =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "plainTypes:lookupPlain");
    CatalogSearchResult noteMatch =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "step template note");

    assertEquals(
        List.of("mutationActionTypes:USE_NESTED"), nested.matches().getFirst().relatedEntryIds());
    assertEquals(
        List.of("inspectionQueryTypes:USE_PLAIN"), plain.matches().getFirst().relatedEntryIds());
    assertEquals("inspectionQueryTypes:USE_PLAIN", noteMatch.matches().getFirst().qualifiedId());
    assertEquals(List.of(), noteMatch.matches().getFirst().relatedEntryIds());
    assertEquals(
        "Unsupported top-level type set: bogusTypeSet",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    GridGrindProtocolCatalogLookupSupport.search(
                        bogusTopLevelCoverageCatalog(), "USE_BOGUS_TOP_LEVEL"))
            .getMessage());
  }

  @Test
  void searchIndexesReferencedTopLevelFamiliesAndCollapsesSupportingGroupEntries() {
    Catalog catalog = topLevelReferenceCoverageCatalog();

    CatalogSearchResult source =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "searchable source");
    CatalogSearchResult persistence =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "searchable persistence");
    CatalogSearchResult step =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "searchable step");
    CatalogSearchResult nestedGroup =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "nested alpha");
    CatalogSearchResult plainGroup =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "plain alpha");

    assertEquals("mutationActionTypes:USE_SOURCE_TYPES", source.matches().getFirst().qualifiedId());
    assertEquals(
        "mutationActionTypes:USE_PERSISTENCE_TYPES",
        persistence.matches().getFirst().qualifiedId());
    assertEquals("mutationActionTypes:USE_STEP_TYPES", step.matches().getFirst().qualifiedId());
    assertTrue(
        nestedGroup.matches().stream()
            .flatMap(match -> match.supportingMatches().stream())
            .map(CatalogSearchMatch::qualifiedId)
            .anyMatch("nestedTypes:lookupNested"::equals));
    assertTrue(
        nestedGroup.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .noneMatch("lookupNested:NESTED_ALPHA"::equals));
    assertTrue(
        plainGroup.matches().stream()
            .flatMap(match -> match.supportingMatches().stream())
            .map(CatalogSearchMatch::qualifiedId)
            .anyMatch("plainTypes:lookupPlain"::equals));
    assertTrue(
        plainGroup.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .noneMatch("lookupPlain:PLAIN_ALPHA"::equals));
  }

  @Test
  void exactSupportingLookupsRemainStandaloneWhileAlsoPointingBackToOwningOperations() {
    Catalog catalog = lookupCoverageCatalog();

    CatalogSearchResult nestedGroup =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "nestedTypes:lookupNested");
    CatalogSearchResult plainGroup =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "plainTypes:lookupPlain");
    CatalogSearchResult nestedEntry =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "lookupNested:NESTED_ALPHA");
    CatalogSearchResult plainEntry =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "lookupPlain:PLAIN_ALPHA");

    assertEquals("nestedTypes:lookupNested", nestedGroup.matches().getFirst().qualifiedId());
    assertEquals(
        List.of("mutationActionTypes:USE_NESTED"),
        nestedGroup.matches().getFirst().relatedEntryIds());
    assertTrue(
        nestedGroup.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .anyMatch("mutationActionTypes:USE_NESTED"::equals));

    assertEquals("plainTypes:lookupPlain", plainGroup.matches().getFirst().qualifiedId());
    assertEquals(
        List.of("inspectionQueryTypes:USE_PLAIN"),
        plainGroup.matches().getFirst().relatedEntryIds());
    assertTrue(
        plainGroup.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .anyMatch("inspectionQueryTypes:USE_PLAIN"::equals));

    assertEquals("lookupNested:NESTED_ALPHA", nestedEntry.matches().getFirst().qualifiedId());
    assertTrue(
        nestedEntry.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .anyMatch("mutationActionTypes:USE_NESTED"::equals));

    assertEquals("lookupPlain:PLAIN_ALPHA", plainEntry.matches().getFirst().qualifiedId());
    assertTrue(
        plainEntry.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .anyMatch("inspectionQueryTypes:USE_PLAIN"::equals));
  }

  @Test
  void searchTraversesRecursiveReferencedShapesAndRejectsBrokenGroupReferences() {
    Catalog catalog = recursiveReferenceCoverageCatalog();

    CatalogSearchResult nestedLeaf =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "nestedTypes:leafGroup");
    CatalogSearchResult plainLeaf =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "plainTypes:plainLeafGroup");

    assertTrue(
        nestedLeaf
            .matches()
            .getFirst()
            .relatedEntryIds()
            .contains("mutationActionTypes:USE_COMPLEX"));
    assertTrue(
        plainLeaf
            .matches()
            .getFirst()
            .relatedEntryIds()
            .contains("mutationActionTypes:USE_COMPLEX"));
    assertEquals(
        "Unknown nested type group: missingNested",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    GridGrindProtocolCatalogLookupSupport.search(
                        brokenNestedCoverageCatalog(), "BROKEN_NESTED"))
            .getMessage());
    assertEquals(
        "Unknown plain type group: missingPlain",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    GridGrindProtocolCatalogLookupSupport.search(
                        brokenPlainCoverageCatalog(), "BROKEN_PLAIN"))
            .getMessage());
  }

  @Test
  void orphanSupportingShapesRemainSearchableWithoutFabricatingOwners() {
    Catalog catalog = orphanSupportingCoverageCatalog();

    CatalogSearchResult nestedGroup =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "nestedTypes:orphanNested");
    CatalogSearchResult plainGroup =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "plainTypes:orphanPlain");
    CatalogSearchResult nestedSummary =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "orphan nested");
    CatalogSearchResult plainSummary =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "orphan plain");
    CatalogSearchResult nestedEntry =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "orphanNested:ORPHAN_NESTED");
    CatalogSearchResult plainEntry =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "orphanPlain:ORPHAN_PLAIN");

    assertEquals("nestedTypes:orphanNested", nestedGroup.matches().getFirst().qualifiedId());
    assertEquals(List.of(), nestedGroup.matches().getFirst().relatedEntryIds());
    assertEquals("plainTypes:orphanPlain", plainGroup.matches().getFirst().qualifiedId());
    assertEquals(List.of(), plainGroup.matches().getFirst().relatedEntryIds());
    assertEquals("nestedTypes:orphanNested", nestedSummary.matches().getFirst().qualifiedId());
    assertEquals("plainTypes:orphanPlain", plainSummary.matches().getFirst().qualifiedId());
    assertEquals("orphanNested:ORPHAN_NESTED", nestedEntry.matches().getFirst().qualifiedId());
    assertEquals("orphanPlain:ORPHAN_PLAIN", plainEntry.matches().getFirst().qualifiedId());
  }

  @Test
  void requestEntriesRemainStandaloneWhenTheyDoNotBelongToAnySupportingGroup() {
    CatalogSearchResult request =
        GridGrindProtocolCatalogLookupSupport.search(lookupCoverageCatalog(), "synthetic request");

    assertEquals("requestType:REQUEST", request.matches().getFirst().qualifiedId());
    assertEquals(List.of(), request.matches().getFirst().relatedEntryIds());
    assertEquals(List.of(), request.matches().getFirst().supportingMatches());
  }

  @Test
  void duplicateSupportingPathsCollapseToOneStandaloneGroupMatchPerLookupId() {
    Catalog catalog = lookupCoverageCatalog();

    CatalogSearchResult nested =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "lookupNested");
    CatalogSearchResult plain =
        GridGrindProtocolCatalogLookupSupport.search(catalog, "lookupPlain");

    assertEquals(
        1L,
        nested.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .filter("nestedTypes:lookupNested"::equals)
            .count());
    assertTrue(
        nested.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .anyMatch("mutationActionTypes:USE_NESTED"::equals));
    assertEquals(
        1L,
        plain.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .filter("plainTypes:lookupPlain"::equals)
            .count());
    assertTrue(
        plain.matches().stream()
            .map(CatalogSearchMatch::qualifiedId)
            .anyMatch("inspectionQueryTypes:USE_PLAIN"::equals));
  }

  @Test
  void groupedSearchFallsThroughMatchingGroupsUntilTheLookupIdAlsoMatches() {
    Catalog catalog = groupLookupBranchCoverageCatalog();
    CatalogSearchMatch supportingEntry =
        new CatalogSearchMatch(
            "lookupNestedOther", "ENTRY_BETA", "lookupNestedOther:ENTRY_BETA", "ENTRY", "summary");

    List<RankedSearchMatch> grouped =
        GridGrindProtocolCatalogLookupSupport.groupSearchMatches(
            catalog, List.of(new RankedSearchMatch(1, supportingEntry)), "not-exact");

    assertEquals("nestedTypes:lookupNestedOther", grouped.getFirst().match().qualifiedId());
  }

  @Test
  void lookupValueGroupTypesPublishOwningTopLevelEntryIds() {
    Catalog catalog = lookupCoverageCatalog();

    var nestedValue =
        new CatalogLookupValue.NestedGroupLookupValue(catalog.nestedTypes().getFirst());
    var plainValue = new CatalogLookupValue.PlainGroupLookupValue(catalog.plainTypes().getFirst());

    assertEquals(List.of("mutationActionTypes:USE_NESTED"), nestedValue.relatedEntryIds(catalog));
    assertEquals(List.of("inspectionQueryTypes:USE_PLAIN"), plainValue.relatedEntryIds(catalog));
  }

  @Test
  void directGroupedSearchSupportCoversDuplicateCollapseAndBrokenOwnerResolution() {
    Catalog catalog = lookupCoverageCatalog();
    CatalogSearchMatch request =
        new CatalogSearchMatch("requestType", "REQUEST", "requestType:REQUEST", "ENTRY", "summary");
    CatalogSearchMatch lowerPriorityRequest =
        new CatalogSearchMatch("requestType", "REQUEST", "requestType:REQUEST", "ENTRY", "summary");
    CatalogSearchMatch exactGroup =
        new CatalogSearchMatch(
            "nestedTypes",
            "lookupNested",
            "nestedTypes:lookupNested",
            "NESTED_GROUP",
            "summary",
            List.of("mutationActionTypes:USE_NESTED"));
    CatalogSearchMatch supportingEntry =
        new CatalogSearchMatch(
            "lookupNested",
            "NESTED_ALPHA",
            "lookupNested:NESTED_ALPHA",
            "ENTRY",
            "summary",
            List.of("mutationActionTypes:USE_NESTED"));
    List<RankedSearchMatch> grouped =
        GridGrindProtocolCatalogLookupSupport.groupSearchMatches(
            catalog,
            List.of(
                new RankedSearchMatch(5, lowerPriorityRequest),
                new RankedSearchMatch(0, exactGroup),
                new RankedSearchMatch(1, supportingEntry),
                new RankedSearchMatch(2, request)),
            "lookupnested");

    assertEquals("nestedTypes:lookupNested", grouped.getFirst().match().qualifiedId());
    assertEquals(
        1L,
        grouped.stream()
            .map(match -> match.match().qualifiedId())
            .filter("nestedTypes:lookupNested"::equals)
            .count());
    assertTrue(
        grouped.stream()
            .map(match -> match.match().qualifiedId())
            .anyMatch("mutationActionTypes:USE_NESTED"::equals));
    assertTrue(
        grouped.stream()
            .map(match -> match.match().qualifiedId())
            .anyMatch("requestType:REQUEST"::equals));

    CatalogSearchMatch malformedSupportingEntry =
        new CatalogSearchMatch(
            "nestedTypes", "ODD_ENTRY", "nestedTypes:ODD_ENTRY", "ENTRY", "summary");
    List<RankedSearchMatch> malformedGrouped =
        GridGrindProtocolCatalogLookupSupport.groupSearchMatches(
            catalog, List.of(new RankedSearchMatch(3, malformedSupportingEntry)), "not-exact");
    assertEquals("nestedTypes:ODD_ENTRY", malformedGrouped.getFirst().match().qualifiedId());

    CatalogSearchMatch brokenSupport =
        new CatalogSearchMatch(
            "nestedTypes",
            "lookupNested",
            "nestedTypes:lookupNested",
            "NESTED_GROUP",
            "summary",
            List.of("lookupNested:NESTED_ALPHA"));
    assertEquals(
        "Catalog related entry id did not resolve: lookupNested:NESTED_ALPHA",
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolCatalogLookupSupport.groupSearchMatches(
                        catalog, List.of(new RankedSearchMatch(0, brokenSupport)), "lookupnested"))
            .getMessage());

    CatalogSearchMatch brokenGroupOwner =
        new CatalogSearchMatch(
            "nestedTypes",
            "lookupNested",
            "nestedTypes:lookupNested",
            "NESTED_GROUP",
            "summary",
            List.of("nestedTypes:lookupNested"));
    assertEquals(
        "Catalog related entry id did not resolve: nestedTypes:lookupNested",
        assertThrows(
                IllegalStateException.class,
                () ->
                    GridGrindProtocolCatalogLookupSupport.groupSearchMatches(
                        catalog,
                        List.of(new RankedSearchMatch(0, brokenGroupOwner)),
                        "lookupnested"))
            .getMessage());
  }

  @Test
  void catalogSearchMatchConvenienceConstructorCopiesRelatedIds() {
    CatalogSearchMatch emptyRelated =
        new CatalogSearchMatch(
            "mutationActionTypes", "SET_CELL", "mutationActionTypes:SET_CELL", "ENTRY", "summary");
    List<String> mutable = new ArrayList<>(List.of("mutationActionTypes:SET_CELL"));
    CatalogSearchMatch copied =
        new CatalogSearchMatch(
            "nestedTypes",
            "cellInputTypes",
            "nestedTypes:cellInputTypes",
            "NESTED_GROUP",
            "summary",
            mutable);
    mutable.add("mutationActionTypes:SET_RANGE");

    assertEquals(List.of(), emptyRelated.relatedEntryIds());
    assertEquals(List.of("mutationActionTypes:SET_CELL"), copied.relatedEntryIds());
    assertThrows(
        UnsupportedOperationException.class,
        () -> copied.relatedEntryIds().add("mutationActionTypes:SET_RANGE"));
  }

  private static Catalog lookupCoverageCatalog() {
    TypeEntry source =
        new TypeEntry("SOURCE_ALPHA", "Source alpha.", List.of(field("path", scalarString())));
    TypeEntry nestedEntry =
        new TypeEntry("NESTED_ALPHA", "Nested alpha.", List.of(field("title", scalarString())));
    TypeEntry plainEntry =
        new TypeEntry("PLAIN_ALPHA", "Plain alpha.", List.of(field("label", scalarString())));
    TypeEntry useNested =
        new TypeEntry(
            "USE_NESTED",
            "Uses nested lookup shapes.",
            List.of(field("nested", new FieldShape.NestedTypeGroupRef("lookupNested"))),
            List.of(),
            Optional.empty(),
            Optional.of(
                new ProtocolStepTemplate(
                    "MUTATION_STEP",
                    "action",
                    objectTemplate("USE_NESTED"),
                    List.of("step template note"))));
    TypeEntry usePlain =
        new TypeEntry(
            "USE_PLAIN",
            "Uses plain lookup shapes.",
            List.of(field("plain", new FieldShape.PlainTypeGroupRef("lookupPlain"))),
            List.of(),
            Optional.empty(),
            Optional.of(
                new ProtocolStepTemplate(
                    "INSPECTION_STEP",
                    "query",
                    objectTemplate("USE_PLAIN"),
                    List.of("step template note"))));
    return new Catalog(
        GridGrindProtocolVersion.current(),
        "type",
        new TypeEntry("REQUEST", "Synthetic request.", List.of()),
        List.of(source),
        List.of(),
        List.of(),
        List.of(useNested),
        List.of(),
        List.of(usePlain),
        List.of(new NestedTypeGroup("lookupNested", "type", List.of(nestedEntry))),
        List.of(new PlainTypeGroup("lookupPlain", plainEntry)));
  }

  private static Catalog bogusTopLevelCoverageCatalog() {
    TypeEntry useBogusTopLevel =
        new TypeEntry(
            "USE_BOGUS_TOP_LEVEL",
            "Unknown top-level set in referenced-shape text.",
            List.of(field("source", new FieldShape.TopLevelTypeSetRef("bogusTypeSet"))));
    return new Catalog(
        GridGrindProtocolVersion.current(),
        "type",
        new TypeEntry("REQUEST", "Synthetic request.", List.of()),
        List.of(),
        List.of(),
        List.of(),
        List.of(useBogusTopLevel),
        List.of(),
        List.of(),
        List.of(),
        List.of());
  }

  private static Catalog topLevelReferenceCoverageCatalog() {
    TypeEntry source =
        new TypeEntry("SOURCE_ALPHA", "Source alpha.", List.of(field("path", scalarString())));
    TypeEntry persistence =
        new TypeEntry(
            "PERSIST_ALPHA", "Persistence alpha.", List.of(field("path", scalarString())));
    TypeEntry step =
        new TypeEntry("STEP_ALPHA", "Step alpha.", List.of(field("stepId", scalarString())));
    TypeEntry nestedEntry =
        new TypeEntry("NESTED_ALPHA", "Nested alpha.", List.of(field("title", scalarString())));
    TypeEntry plainEntry =
        new TypeEntry("PLAIN_ALPHA", "Plain alpha.", List.of(field("label", scalarString())));
    TypeEntry useSource =
        new TypeEntry(
            "USE_SOURCE_TYPES",
            "Searchable through source type references.",
            List.of(field("source", new FieldShape.TopLevelTypeSetRef("sourceTypes"))));
    TypeEntry usePersistence =
        new TypeEntry(
            "USE_PERSISTENCE_TYPES",
            "Searchable through persistence type references.",
            List.of(field("persistence", new FieldShape.TopLevelTypeSetRef("persistenceTypes"))));
    TypeEntry useStep =
        new TypeEntry(
            "USE_STEP_TYPES",
            "Searchable through step type references.",
            List.of(field("step", new FieldShape.TopLevelTypeSetRef("stepTypes"))));
    TypeEntry useNested =
        new TypeEntry(
            "USE_NESTED_GROUP",
            "Searchable through nested group references.",
            List.of(field("nested", new FieldShape.NestedTypeGroupRef("lookupNested"))));
    TypeEntry usePlain =
        new TypeEntry(
            "USE_PLAIN_GROUP",
            "Searchable through plain group references.",
            List.of(field("plain", new FieldShape.PlainTypeGroupRef("lookupPlain"))));
    return new Catalog(
        GridGrindProtocolVersion.current(),
        "type",
        new TypeEntry("REQUEST", "Synthetic request.", List.of()),
        List.of(source),
        List.of(persistence),
        List.of(step),
        List.of(useSource, usePersistence, useStep, useNested, usePlain),
        List.of(),
        List.of(),
        List.of(new NestedTypeGroup("lookupNested", "type", List.of(nestedEntry))),
        List.of(new PlainTypeGroup("lookupPlain", plainEntry)));
  }

  private static Catalog recursiveReferenceCoverageCatalog() {
    TypeEntry recursiveNested =
        new TypeEntry(
            "RECURSIVE_NESTED",
            "Recursive nested entry.",
            List.of(field("self", new FieldShape.NestedTypeGroupRef("recursiveGroup"))));
    TypeEntry nestedLeaf =
        new TypeEntry("LEAF_ENTRY", "Leaf nested entry.", List.of(field("title", scalarString())));
    TypeEntry unionEntry =
        new TypeEntry(
            "UNION_ENTRY",
            "Union nested entry.",
            List.of(field("leaf", new FieldShape.NestedTypeGroupRef("leafGroup"))));
    TypeEntry recursivePlain =
        new TypeEntry(
            "RECURSIVE_PLAIN",
            "Recursive plain entry.",
            List.of(field("self", new FieldShape.PlainTypeGroupRef("recursivePlainGroup"))));
    TypeEntry plainLeaf =
        new TypeEntry("PLAIN_LEAF", "Plain leaf entry.", List.of(field("label", scalarString())));
    TypeEntry action =
        new TypeEntry(
            "USE_COMPLEX",
            "Complex reference coverage entry.",
            List.of(
                field("recursiveProbe", new FieldShape.NestedTypeGroupRef("recursiveGroup")),
                field("unionProbe", new FieldShape.NestedTypeGroupUnionRef(List.of("unionGroup"))),
                field("leaf", new FieldShape.NestedTypeGroupRef("leafGroup")),
                field("recursivePlain", new FieldShape.PlainTypeGroupRef("recursivePlainGroup")),
                field("plainLeaf", new FieldShape.PlainTypeGroupRef("plainLeafGroup"))));
    return new Catalog(
        GridGrindProtocolVersion.current(),
        "type",
        new TypeEntry("REQUEST", "Synthetic request.", List.of()),
        List.of(),
        List.of(),
        List.of(),
        List.of(action),
        List.of(),
        List.of(),
        List.of(
            new NestedTypeGroup("recursiveGroup", "type", List.of(recursiveNested)),
            new NestedTypeGroup("leafGroup", "type", List.of(nestedLeaf)),
            new NestedTypeGroup("unionGroup", "type", List.of(unionEntry))),
        List.of(
            new PlainTypeGroup("recursivePlainGroup", recursivePlain),
            new PlainTypeGroup("plainLeafGroup", plainLeaf)));
  }

  private static Catalog brokenNestedCoverageCatalog() {
    TypeEntry action =
        new TypeEntry(
            "BROKEN_NESTED",
            "Broken nested group reference.",
            List.of(field("nested", new FieldShape.NestedTypeGroupRef("missingNested"))));
    return new Catalog(
        GridGrindProtocolVersion.current(),
        "type",
        new TypeEntry("REQUEST", "Synthetic request.", List.of()),
        List.of(),
        List.of(),
        List.of(),
        List.of(action),
        List.of(),
        List.of(),
        List.of(),
        List.of());
  }

  private static Catalog brokenPlainCoverageCatalog() {
    TypeEntry action =
        new TypeEntry(
            "BROKEN_PLAIN",
            "Broken plain group reference.",
            List.of(field("plain", new FieldShape.PlainTypeGroupRef("missingPlain"))));
    return new Catalog(
        GridGrindProtocolVersion.current(),
        "type",
        new TypeEntry("REQUEST", "Synthetic request.", List.of()),
        List.of(),
        List.of(),
        List.of(),
        List.of(action),
        List.of(),
        List.of(),
        List.of(),
        List.of());
  }

  private static Catalog orphanSupportingCoverageCatalog() {
    TypeEntry nestedEntry =
        new TypeEntry(
            "ORPHAN_NESTED", "Orphan nested entry.", List.of(field("title", scalarString())));
    TypeEntry plainEntry =
        new TypeEntry(
            "ORPHAN_PLAIN", "Orphan plain entry.", List.of(field("label", scalarString())));
    return new Catalog(
        GridGrindProtocolVersion.current(),
        "type",
        new TypeEntry("REQUEST", "Synthetic request.", List.of()),
        List.of(),
        List.of(),
        List.of(),
        List.of(),
        List.of(),
        List.of(),
        List.of(new NestedTypeGroup("orphanNested", "type", List.of(nestedEntry))),
        List.of(new PlainTypeGroup("orphanPlain", plainEntry)));
  }

  private static Catalog groupLookupBranchCoverageCatalog() {
    TypeEntry firstNested =
        new TypeEntry(
            "ENTRY_ALPHA", "First nested entry.", List.of(field("title", scalarString())));
    TypeEntry secondNested =
        new TypeEntry(
            "ENTRY_BETA", "Second nested entry.", List.of(field("title", scalarString())));
    return new Catalog(
        GridGrindProtocolVersion.current(),
        "type",
        new TypeEntry("REQUEST", "Synthetic request.", List.of()),
        List.of(),
        List.of(),
        List.of(),
        List.of(),
        List.of(),
        List.of(),
        List.of(
            new NestedTypeGroup("lookupNested", "type", List.of(firstNested)),
            new NestedTypeGroup("lookupNestedOther", "type", List.of(secondNested))),
        List.of());
  }

  private static FieldEntry field(String name, FieldShape shape) {
    return new FieldEntry(name, FieldRequirement.REQUIRED, shape, List.of());
  }

  private static FieldShape.Scalar scalarString() {
    return new FieldShape.Scalar(ScalarType.STRING);
  }

  private static tools.jackson.databind.node.ObjectNode objectTemplate(String typeId) {
    tools.jackson.databind.node.ObjectNode node =
        tools.jackson.databind.node.JsonNodeFactory.instance.objectNode();
    node.put("type", typeId);
    return node;
  }
}
