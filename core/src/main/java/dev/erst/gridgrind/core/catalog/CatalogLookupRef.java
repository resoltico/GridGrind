package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.vocabulary.TextValidation;

/** One published lookup surface entry: catalog group, lookup id, qualified id, and its value. */
record CatalogLookupRef(
    String catalogGroup, String lookupId, String qualifiedId, CatalogLookupValue value) {
  CatalogLookupRef {
    catalogGroup = TextValidation.requireNonBlank(catalogGroup, "catalogGroup");
    lookupId = TextValidation.requireNonBlank(lookupId, "lookupId");
    qualifiedId = TextValidation.requireNonBlank(qualifiedId, "qualifiedId");
    java.util.Objects.requireNonNull(value, "value must not be null");
  }
}
