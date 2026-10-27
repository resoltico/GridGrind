package dev.erst.gridgrind.cli.render;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.core.discovery.DiscoveryValidation;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
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
    catalogGroup = TextValidation.requireNonBlank(catalogGroup, "catalogGroup");
    lookupId = TextValidation.requireNonBlank(lookupId, "lookupId");
    qualifiedId = TextValidation.requireNonBlank(qualifiedId, "qualifiedId");
    kind = TextValidation.requireNonBlank(kind, "kind");
    summary = TextValidation.requireNonBlank(summary, "summary");
    relatedEntryIds =
        DiscoveryValidation.copyOptionalStringsAllowEmpty(relatedEntryIds, "relatedEntryIds");
    supportingQualifiedIds =
        DiscoveryValidation.copyOptionalStringsAllowEmpty(
            supportingQualifiedIds, "supportingQualifiedIds");
  }
}
