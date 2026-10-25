package dev.erst.gridgrind.contract.catalog;

import java.util.List;

/** Internal descriptor for one field whose presence is gated by projected read facets. */
record CatalogProjectedField(String name, List<String> projectedByFacets) {
  CatalogProjectedField(String name, String... projectedByFacets) {
    this(name, List.of(projectedByFacets));
  }

  CatalogProjectedField {
    name = CatalogRecordValidation.requireNonBlank(name, "name");
    projectedByFacets =
        CatalogRecordValidation.copyUniqueStrings(projectedByFacets, "projectedByFacets");
    if (projectedByFacets.isEmpty()) {
      throw new IllegalArgumentException("projectedByFacets must not be empty");
    }
  }
}
