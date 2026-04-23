package dev.erst.gridgrind.contract.catalog;

import java.util.List;
import java.util.Objects;

/** Case-insensitive protocol-catalog search result returned by CLI discovery commands. */
public record CatalogSearchResult(String query, List<CatalogSearchMatch> matches) {
  public CatalogSearchResult {
    query = CatalogRecordValidation.requireNonBlank(query, "query");
    Objects.requireNonNull(matches, "matches must not be null");
    for (CatalogSearchMatch match : matches) {
      if (match == null) {
        throw new IllegalArgumentException("matches must not contain nulls");
      }
    }
    matches = List.copyOf(matches);
  }
}
