package dev.erst.gridgrind.contract.catalog;

/** One published lookup surface entry: catalog group, lookup id, qualified id, and its value. */
record CatalogLookupRef(
    String catalogGroup, String lookupId, String qualifiedId, CatalogLookupValue value) {
  CatalogLookupRef {
    catalogGroup = CatalogRecordValidation.requireNonBlank(catalogGroup, "catalogGroup");
    lookupId = CatalogRecordValidation.requireNonBlank(lookupId, "lookupId");
    qualifiedId = CatalogRecordValidation.requireNonBlank(qualifiedId, "qualifiedId");
    java.util.Objects.requireNonNull(value, "value must not be null");
  }
}
