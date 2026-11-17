package dev.erst.gridgrind.core.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertSame;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Checks closed lookup results reuse published catalog records and semantic metadata. */
class CatalogLookupPublicContractTest {
  @Test
  void lookupReturnsThePublishedRecordsThroughAnExhaustivePublicFamily() {
    var catalog = GridGrindProtocolCatalog.catalog();
    List<CatalogLookupValue> values =
        List.of(
            GridGrindProtocolCatalog.lookupValueFor("requestType:WorkbookPlan").orElseThrow(),
            GridGrindProtocolCatalog.lookupValueFor(
                    "nestedTypes:" + catalog.nestedTypes().getFirst().group())
                .orElseThrow(),
            GridGrindProtocolCatalog.lookupValueFor(
                    "plainTypes:" + catalog.plainTypes().getFirst().group())
                .orElseThrow(),
            GridGrindProtocolCatalog.lookupValueFor("mutationActionTypes").orElseThrow());
    assertEquals(
        List.of("ENTRY", "NESTED", "PLAIN", "TOP_LEVEL"),
        values.stream()
            .map(value ->
                switch (value) {
                  case TypeEntry ignored -> "ENTRY";
                  case NestedTypeGroup ignored -> "NESTED";
                  case PlainTypeGroup ignored -> "PLAIN";
                  case TopLevelTypeGroup ignored -> "TOP_LEVEL";
                })
            .toList());
    assertSame(catalog.requestType(), values.get(0));
    assertSame(catalog.nestedTypes().getFirst(), values.get(1));
    assertSame(catalog.plainTypes().getFirst(), values.get(2));
    assertEquals(
        WorkbookPlan.class.getAnnotation(CatalogType.class).summary(),
        catalog.requestType().summary());
  }
}
