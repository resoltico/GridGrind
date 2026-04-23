package dev.erst.gridgrind.contract.catalog;

/** One protocol-catalog search hit suitable for CLI and agent-facing discovery. */
public record CatalogSearchMatch(
    String catalogGroup, String lookupId, String qualifiedId, String kind, String summary) {
  public CatalogSearchMatch {
    catalogGroup = CatalogRecordValidation.requireNonBlank(catalogGroup, "catalogGroup");
    lookupId = CatalogRecordValidation.requireNonBlank(lookupId, "lookupId");
    qualifiedId = CatalogRecordValidation.requireNonBlank(qualifiedId, "qualifiedId");
    kind = CatalogRecordValidation.requireNonBlank(kind, "kind");
    summary = CatalogRecordValidation.requireNonBlank(summary, "summary");
  }
}
