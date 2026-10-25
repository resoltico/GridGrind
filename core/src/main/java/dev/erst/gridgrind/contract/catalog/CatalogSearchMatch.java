package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.List;

/** One protocol-catalog search hit suitable for CLI and agent-facing discovery. */
public record CatalogSearchMatch(
    String catalogGroup,
    String lookupId,
    String qualifiedId,
    String kind,
    String summary,
    @JsonInclude(JsonInclude.Include.NON_ABSENT)
        java.util.Optional<ProtocolStepTemplate> stepTemplate,
    @JsonInclude(JsonInclude.Include.NON_EMPTY) List<String> relatedEntryIds,
    @JsonInclude(JsonInclude.Include.NON_EMPTY) List<CatalogSearchMatch> supportingMatches) {
  /** Creates one search hit with no related top-level entry ids. */
  public CatalogSearchMatch(
      String catalogGroup, String lookupId, String qualifiedId, String kind, String summary) {
    this(
        catalogGroup,
        lookupId,
        qualifiedId,
        kind,
        summary,
        java.util.Optional.empty(),
        List.of(),
        List.of());
  }

  /** Creates one search hit with related top-level entry ids but without grouped support hits. */
  public CatalogSearchMatch(
      String catalogGroup,
      String lookupId,
      String qualifiedId,
      String kind,
      String summary,
      List<String> relatedEntryIds) {
    this(
        catalogGroup,
        lookupId,
        qualifiedId,
        kind,
        summary,
        java.util.Optional.empty(),
        relatedEntryIds,
        List.of());
  }

  public CatalogSearchMatch {
    catalogGroup = CatalogRecordValidation.requireNonBlank(catalogGroup, "catalogGroup");
    lookupId = CatalogRecordValidation.requireNonBlank(lookupId, "lookupId");
    qualifiedId = CatalogRecordValidation.requireNonBlank(qualifiedId, "qualifiedId");
    kind = CatalogRecordValidation.requireNonBlank(kind, "kind");
    summary = CatalogRecordValidation.requireNonBlank(summary, "summary");
    java.util.Objects.requireNonNull(stepTemplate, "stepTemplate must not be null");
    relatedEntryIds = List.copyOf(relatedEntryIds);
    java.util.Objects.requireNonNull(supportingMatches, "supportingMatches must not be null");
    supportingMatches = List.copyOf(supportingMatches);
  }
}
