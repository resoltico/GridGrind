package dev.erst.gridgrind.cli.discovery;

import com.fasterxml.jackson.annotation.JsonInclude;
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
    catalogGroup = CliDiscoveryValidation.requireNonBlank(catalogGroup, "catalogGroup");
    lookupId = CliDiscoveryValidation.requireNonBlank(lookupId, "lookupId");
    qualifiedId = CliDiscoveryValidation.requireNonBlank(qualifiedId, "qualifiedId");
    kind = CliDiscoveryValidation.requireNonBlank(kind, "kind");
    summary = CliDiscoveryValidation.requireNonBlank(summary, "summary");
    relatedEntryIds =
        CliDiscoveryValidation.copyOptionalStringsAllowEmpty(relatedEntryIds, "relatedEntryIds");
    supportingQualifiedIds =
        CliDiscoveryValidation.copyOptionalStringsAllowEmpty(
            supportingQualifiedIds, "supportingQualifiedIds");
  }
}
