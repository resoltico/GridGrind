package dev.erst.gridgrind.contract.catalog;

import java.util.Objects;

/** A scored catalog search match used during search result aggregation. */
record RankedSearchMatch(int rank, CatalogSearchMatch match) {
  RankedSearchMatch {
    Objects.requireNonNull(match, "match must not be null");
  }
}
