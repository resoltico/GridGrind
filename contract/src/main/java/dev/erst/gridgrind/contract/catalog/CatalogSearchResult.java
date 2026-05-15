package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonPropertyOrder;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.List;
import java.util.Objects;

/** Case-insensitive protocol-catalog search result returned by CLI discovery commands. */
@JsonPropertyOrder({"protocolVersion", "query", "totalCount", "matches"})
public record CatalogSearchResult(
    GridGrindProtocolVersion protocolVersion, String query, List<CatalogSearchMatch> matches) {
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
  }

  /** Total number of matches; always equals {@code matches().size()}. */
  @JsonProperty
  public int totalCount() {
    return matches.size();
  }
}
