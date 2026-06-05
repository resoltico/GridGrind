package dev.erst.gridgrind.cli.discovery;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;
import com.fasterxml.jackson.annotation.JsonProperty;
import com.fasterxml.jackson.annotation.JsonPropertyOrder;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import java.util.List;
import java.util.Objects;

/** Summary-first protocol-catalog search result for CLI discovery. */
@JsonIgnoreProperties(ignoreUnknown = true)
@JsonPropertyOrder({"protocolVersion", "query", "totalCount", "matches"})
public record ProtocolCatalogSearchReport(
    GridGrindProtocolVersion protocolVersion,
    String query,
    List<ProtocolCatalogSearchHit> matches) {
  public ProtocolCatalogSearchReport {
    Objects.requireNonNull(protocolVersion, "protocolVersion must not be null");
    query = CliDiscoveryValidation.requireNonBlank(query, "query");
    Objects.requireNonNull(matches, "matches must not be null");
    matches =
        matches.stream()
            .map(match -> Objects.requireNonNull(match, "matches must not contain nulls"))
            .toList();
  }

  /** Total number of matches; always equals {@code matches().size()}. */
  @JsonProperty
  public int totalCount() {
    return matches.size();
  }
}
