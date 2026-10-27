package dev.erst.gridgrind.cli.render;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.core.discovery.DiscoveryValidation;
import java.util.List;

/** Summary-first search hit published by the CLI protocol-catalog discovery surface. */
public record ProtocolCatalogSearchHit(
    String catalogGroup,
    String lookupId,
    String qualifiedId,
    String kind,
    String summary,
    @JsonInclude(JsonInclude.Include.NON_EMPTY) List<String> relatedEntryIds,
    @JsonInclude(JsonInclude.Include.NON_EMPTY) List<String> supportingQualifiedIds) {
  public ProtocolCatalogSearchHit {
    catalogGroup = DiscoveryValidation.requireNonBlank(catalogGroup, "catalogGroup");
    lookupId = DiscoveryValidation.requireNonBlank(lookupId, "lookupId");
    qualifiedId = DiscoveryValidation.requireNonBlank(qualifiedId, "qualifiedId");
    kind = DiscoveryValidation.requireNonBlank(kind, "kind");
    summary = DiscoveryValidation.requireNonBlank(summary, "summary");
    relatedEntryIds =
        DiscoveryValidation.copyOptionalStringsAllowEmpty(relatedEntryIds, "relatedEntryIds");
    supportingQualifiedIds =
        DiscoveryValidation.copyOptionalStringsAllowEmpty(
            supportingQualifiedIds, "supportingQualifiedIds");
  }
}
