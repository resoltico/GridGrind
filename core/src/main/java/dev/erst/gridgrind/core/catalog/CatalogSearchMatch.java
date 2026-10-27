package dev.erst.gridgrind.core.catalog;

import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;

/** One protocol-catalog search hit suitable for CLI and agent-facing discovery. */
public record CatalogSearchMatch(
    String catalogGroup,
    String lookupId,
    String qualifiedId,
    String kind,
    String summary,
    java.util.Optional<ProtocolStepTemplate> stepTemplate,
    @WireDefault(value = "[]", omit = true) List<String> relatedEntryIds,
    @WireDefault(value = "[]", omit = true) List<CatalogSearchMatch> supportingMatches) {
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
    catalogGroup = TextValidation.requireNonBlank(catalogGroup, "catalogGroup");
    lookupId = TextValidation.requireNonBlank(lookupId, "lookupId");
    qualifiedId = TextValidation.requireNonBlank(qualifiedId, "qualifiedId");
    kind = TextValidation.requireNonBlank(kind, "kind");
    summary = TextValidation.requireNonBlank(summary, "summary");
    java.util.Objects.requireNonNull(stepTemplate, "stepTemplate must not be null");
    relatedEntryIds = List.copyOf(relatedEntryIds);
    java.util.Objects.requireNonNull(supportingMatches, "supportingMatches must not be null");
    supportingMatches = List.copyOf(supportingMatches);
  }
}
