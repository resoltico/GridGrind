package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.plan.GridGrindProtocolVersion;
import java.util.List;
import java.util.Objects;

/** Case-insensitive protocol-catalog search result returned by CLI discovery commands. */
public record CatalogSearchResult(
    GridGrindProtocolVersion protocolVersion,
    String query,
    int totalCount,
    List<CatalogSearchMatch> matches) {
  public CatalogSearchResult {
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    query = CatalogRecordValidation.requireNonBlank(query, "query");
    Objects.requireNonNull(matches, "matches must not be null");
    for (CatalogSearchMatch match : matches) {
      if (match == null) {
        throw new IllegalArgumentException("matches must not contain nulls");
      }
    }
    matches = List.copyOf(matches);
    if (totalCount != matches.size()) {
      throw new IllegalArgumentException("totalCount must equal matches.size()");
    }
  }

  /** Computes the match count for a search assembled by Java callers. */
  public CatalogSearchResult(
      GridGrindProtocolVersion protocolVersion, String query, List<CatalogSearchMatch> matches) {
    this(protocolVersion, query, matches.size(), matches);
  }
}
