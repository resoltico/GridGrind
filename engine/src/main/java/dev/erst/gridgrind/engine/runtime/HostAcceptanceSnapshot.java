package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;

/** Immutable before-mutation facts retained only for one host-required preservation comparison. */
record HostAcceptanceSnapshot(
    List<SemanticPreservation> semanticPreservation,
    List<OpaquePartPreservation> opaquePartPreservation) {
  HostAcceptanceSnapshot {
    semanticPreservation = List.copyOf(semanticPreservation);
    opaquePartPreservation = List.copyOf(opaquePartPreservation);
  }

  static HostAcceptanceSnapshot empty() {
    return new HostAcceptanceSnapshot(List.of(), List.of());
  }

  Path sourceFor(GridGrindHostAcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts requirement) {
    Objects.requireNonNull(requirement, "requirement must not be null");
    return opaquePartPreservation.stream()
        .filter(candidate -> candidate.requirement().equals(requirement))
        .findFirst()
        .map(OpaquePartPreservation::sourceArtifact)
        .orElseThrow(
            () ->
                new IllegalStateException(
                    "host opaque-part preservation was not captured before mutation"));
  }

  List<InspectionResult> baselineFor(
      GridGrindHostAcceptancePolicy.Requirement.PreserveInspectionFacts requirement) {
    Objects.requireNonNull(requirement, "requirement must not be null");
    return semanticPreservation.stream()
        .filter(candidate -> candidate.requirement().equals(requirement))
        .findFirst()
        .map(SemanticPreservation::baseline)
        .orElseThrow(
            () ->
                new IllegalStateException(
                    "host semantic-preservation requirement was not captured before mutation"));
  }

  record SemanticPreservation(
      GridGrindHostAcceptancePolicy.Requirement.PreserveInspectionFacts requirement,
      List<InspectionResult> baseline) {
    SemanticPreservation {
      Objects.requireNonNull(requirement, "requirement must not be null");
      baseline = List.copyOf(baseline);
      if (baseline.size() != requirement.inspections().size()) {
        throw new IllegalArgumentException("baseline size must match preservation inspections");
      }
    }
  }

  /** One source package retained only while an admitted execution needs byte-preservation proof. */
  record OpaquePartPreservation(
      GridGrindHostAcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts requirement,
      Path sourceArtifact) {
    OpaquePartPreservation {
      Objects.requireNonNull(requirement, "requirement must not be null");
      Objects.requireNonNull(sourceArtifact, "sourceArtifact must not be null");
    }
  }
}
