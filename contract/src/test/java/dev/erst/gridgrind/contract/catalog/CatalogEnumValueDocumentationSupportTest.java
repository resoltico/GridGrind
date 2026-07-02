package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.query.CellReadFacet;
import dev.erst.gridgrind.contract.query.CellReadProjection;
import java.lang.reflect.RecordComponent;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Focused coverage for machine-readable enum value documentation helpers. */
class CatalogEnumValueDocumentationSupportTest {
  @Test
  void enumValueDocsHandleDirectOptionalListAndUnsupportedTypes() throws Exception {
    assertEquals(
        "VALUE",
        CatalogEnumValueDocumentationSupport.enumValueDocs(CellReadFacet.class).getFirst().value());
    assertEquals(
        "VALUE",
        CatalogEnumValueDocumentationSupport.enumValueDocs(
                recordComponent(EnumDocFixture.class, "optionalFacet").getGenericType())
            .getFirst()
            .value());
    assertEquals(
        "VALUE",
        CatalogEnumValueDocumentationSupport.enumValueDocs(
                recordComponent(CellReadProjection.class, "facets").getGenericType())
            .getFirst()
            .value());
    assertEquals(
        List.of(),
        CatalogEnumValueDocumentationSupport.enumValueDocs(
            recordComponent(EnumDocFixture.class, "cells").getGenericType()));
  }

  @Test
  void orderedEnumDocsRejectIncompleteCoverage() {
    IllegalStateException missingDoc =
        assertThrows(
            IllegalStateException.class,
            () ->
                CatalogEnumValueDocumentationSupport.orderedEnumDocs(
                    CellReadFacet.class, Map.of("VALUE", "Only one token is documented.")));

    assertEquals(
        "Enum value docs must cover every published token for " + CellReadFacet.class.getName(),
        missingDoc.getMessage());
  }

  private static RecordComponent recordComponent(Class<? extends Record> recordType, String name) {
    for (RecordComponent component : recordType.getRecordComponents()) {
      if (component.getName().equals(name)) {
        return component;
      }
    }
    throw new IllegalArgumentException("Unknown record component: " + name);
  }

  private record EnumDocFixture(
      Optional<CellReadFacet> optionalFacet, Map<String, CellInput> cells) {}
}
