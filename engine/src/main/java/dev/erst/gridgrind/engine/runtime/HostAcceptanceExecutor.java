package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.grant.AcceptancePolicy;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.result.WorkbookExecutionEvidence;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookLocation;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Evaluates host-owned terminal acceptance after mutations and before final publication. */
final class HostAcceptanceExecutor {
  private final StepExecution stepSupport;

  HostAcceptanceExecutor(StepExecution stepSupport) {
    this.stepSupport = Objects.requireNonNull(stepSupport, "stepSupport must not be null");
  }

  HostAcceptanceSnapshot captureBeforeMutation(
      AcceptancePolicy authoredPolicy,
      AcceptancePolicy policy,
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation,
      @Nullable Path materializedSourceArtifact)
      throws IOException {
    Objects.requireNonNull(workbook, "workbook must not be null");
    requirePolicyBinding(authoredPolicy, policy);
    return switch (Objects.requireNonNull(policy, "policy must not be null")) {
      case AcceptancePolicy.MinimumOnly _ -> HostAcceptanceSnapshot.empty();
      case AcceptancePolicy.RequireAll requireAll ->
          captureWorkbookRequirements(
              requireAllPolicy(authoredPolicy, policy),
              requireAll,
              workbook,
              workbookLocation,
              materializedSourceArtifact);
    };
  }

  HostAcceptanceVerification verifyWorkbook(
      AcceptancePolicy authoredPolicy,
      AcceptancePolicy policy,
      HostAcceptanceSnapshot snapshot,
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation,
      List<AssertionResult> assertions,
      List<WorkbookExecutionEvidence.Preservation> preservation)
      throws IOException, AssertionFailedException, PreservationFailedException {
    requirePolicyBinding(authoredPolicy, policy);
    return switch (Objects.requireNonNull(policy, "policy must not be null")) {
      case AcceptancePolicy.MinimumOnly _ -> verifyMinimumOnly();
      case AcceptancePolicy.RequireAll requireAll ->
          verifyRequiredWorkbook(
              requireAllPolicy(authoredPolicy, policy),
              requireAll,
              snapshot,
              workbook,
              workbookLocation,
              assertions,
              preservation);
    };
  }

  /** Establishes any host-required opaque-part byte identity against a verified staged package. */
  List<WorkbookExecutionEvidence.Preservation> verifyStagedArtifact(
      AcceptancePolicy policy,
      HostAcceptanceSnapshot snapshot,
      Path stagedWorkbookArtifact,
      List<WorkbookExecutionEvidence.Preservation> preservation)
      throws IOException, PreservationFailedException {
    Objects.requireNonNull(snapshot, "snapshot must not be null");
    Objects.requireNonNull(stagedWorkbookArtifact, "stagedWorkbookArtifact must not be null");
    return switch (Objects.requireNonNull(policy, "policy must not be null")) {
      case AcceptancePolicy.MinimumOnly _ -> List.of();
      case AcceptancePolicy.RequireAll requireAll ->
          verifyStagedOpaquePartPreservation(
              requireAll, snapshot, stagedWorkbookArtifact, preservation);
    };
  }

  private HostAcceptanceVerification verifyRequiredWorkbook(
      AcceptancePolicy.RequireAll authoredPolicy,
      AcceptancePolicy.RequireAll requireAll,
      HostAcceptanceSnapshot snapshot,
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation,
      List<AssertionResult> assertions,
      List<WorkbookExecutionEvidence.Preservation> preservation)
      throws IOException, AssertionFailedException, PreservationFailedException {
    Objects.requireNonNull(assertions, "assertions must not be null");
    Objects.requireNonNull(preservation, "preservation must not be null");
    for (int requirementIndex = 0;
        requirementIndex < requireAll.requirements().size();
        requirementIndex++) {
      var requirement = requireAll.requirements().get(requirementIndex);
      var original = originalRequirement(authoredPolicy, requireAll, requirementIndex);
      switch (requirement) {
        case AcceptancePolicy.Requirement.TerminalAssertions terminalAssertions -> {
          var originalAssertions =
              ((AcceptancePolicy.Requirement.TerminalAssertions) original).assertions();
          for (int index = 0; index < terminalAssertions.assertions().size(); index++) {
            var assertion = terminalAssertions.assertions().get(index);
            assertions.add(
                stepSupport.executeAssertionStep(
                    (dev.erst.gridgrind.core.plan.PreparedAssertionStep)
                        dev.erst.gridgrind.core.plan.PreparedStep.prepare(
                            originalAssertions.get(index), assertion),
                    workbook,
                    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null"),
                    new ExecutionModeSpec.FullXssf()));
          }
        }
        case AcceptancePolicy.Requirement.PreserveInspectionFacts facts ->
            verifyWorkbookSemanticPreservation(
                (AcceptancePolicy.Requirement.PreserveInspectionFacts) original,
                facts,
                snapshot,
                workbook,
                workbookLocation,
                preservation);
        case AcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts _ -> {}
        case AcceptancePolicy.Requirement.RequireCalculation _ -> {}
      }
    }
    return new HostAcceptanceVerification(assertions, preservation);
  }

  /**
   * Verifies the terminal-only acceptance subset available after mode-policy validation for an
   * event-read or streaming-write execution.
   */
  HostAcceptanceVerification verifyMaterializedTerminalAcceptance(
      AcceptancePolicy authoredPolicy,
      AcceptancePolicy policy,
      Path workbookPath,
      WorkbookLocation workbookLocation,
      List<AssertionResult> assertions)
      throws IOException, AssertionFailedException {
    requirePolicyBinding(authoredPolicy, policy);
    Objects.requireNonNull(assertions, "assertions must not be null");
    if (policy instanceof AcceptancePolicy.RequireAll requireAll
        && requireAll.requirements().stream()
            .anyMatch(AcceptancePolicy.Requirement.TerminalAssertions.class::isInstance)) {
      // Every terminal requirement observes the same completed artifact.
      try (var view =
          stepSupport.openReadView(
              workbookPath,
              Objects.requireNonNull(workbookLocation, "workbookLocation must not be null"))) {
        var originals = requireAllPolicy(authoredPolicy, policy);
        for (int requirementIndex = 0;
            requirementIndex < requireAll.requirements().size();
            requirementIndex++) {
          var requirement = requireAll.requirements().get(requirementIndex);
          var original = originalRequirement(originals, requireAll, requirementIndex);
          if (requirement instanceof AcceptancePolicy.Requirement.TerminalAssertions terminal) {
            var originalAssertions =
                ((AcceptancePolicy.Requirement.TerminalAssertions) original).assertions();
            for (int index = 0; index < terminal.assertions().size(); index++) {
              var prepared =
                  (dev.erst.gridgrind.core.plan.PreparedAssertionStep)
                      dev.erst.gridgrind.core.plan.PreparedStep.prepare(
                          originalAssertions.get(index), terminal.assertions().get(index));
              assertions.add(view.assertThat(prepared));
            }
          }
        }
      }
    }
    return new HostAcceptanceVerification(assertions, List.of());
  }

  private static void requirePolicyBinding(AcceptancePolicy authored, AcceptancePolicy bound) {
    Objects.requireNonNull(authored, "authored policy must not be null");
    Objects.requireNonNull(bound, "bound policy must not be null");
    if (authored.getClass() != bound.getClass())
      throw new IllegalArgumentException("Bound host policy must retain its authored kind");
    if (bound instanceof AcceptancePolicy.RequireAll prepared) {
      var original = requireAllPolicy(authored, bound);
      for (int index = 0; index < prepared.requirements().size(); index++)
        originalRequirement(original, prepared, index);
    }
  }

  private static AcceptancePolicy.RequireAll requireAllPolicy(
      AcceptancePolicy authored, AcceptancePolicy bound) {
    if (!(authored instanceof AcceptancePolicy.RequireAll original)
        || !(bound instanceof AcceptancePolicy.RequireAll prepared)
        || original.requirements().size() != prepared.requirements().size()) {
      throw new IllegalArgumentException(
          "Bound host requirements must retain their authored structure");
    }
    return original;
  }

  private static AcceptancePolicy.Requirement originalRequirement(
      AcceptancePolicy.RequireAll authored, AcceptancePolicy.RequireAll bound, int index) {
    var original = authored.requirements().get(index);
    var prepared = bound.requirements().get(index);
    if (original.getClass() != prepared.getClass())
      throw new IllegalArgumentException("Bound host requirement kind must match its author");
    switch (prepared) {
      case AcceptancePolicy.Requirement.TerminalAssertions terminal -> {
        var originalSteps =
            ((AcceptancePolicy.Requirement.TerminalAssertions) original).assertions();
        requireSameStepCount(originalSteps.size(), terminal.assertions().size());
        for (int step = 0; step < originalSteps.size(); step++)
          dev.erst.gridgrind.core.plan.PreparedStep.prepare(
              originalSteps.get(step), terminal.assertions().get(step));
      }
      case AcceptancePolicy.Requirement.PreserveInspectionFacts facts -> {
        var originalSteps =
            ((AcceptancePolicy.Requirement.PreserveInspectionFacts) original).inspections();
        requireSameStepCount(originalSteps.size(), facts.inspections().size());
        for (int step = 0; step < originalSteps.size(); step++)
          dev.erst.gridgrind.core.plan.PreparedStep.prepare(
              originalSteps.get(step), facts.inspections().get(step));
      }
      case AcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts ignored -> {
        if (!original.equals(prepared))
          throw new IllegalArgumentException("Bound host requirement must match its author");
      }
      case AcceptancePolicy.Requirement.RequireCalculation ignored -> {
        if (!original.equals(prepared))
          throw new IllegalArgumentException("Bound host requirement must match its author");
      }
    }
    return original;
  }

  private static void requireSameStepCount(int authored, int bound) {
    if (authored != bound)
      throw new IllegalArgumentException(
          "Bound host checks must retain their authored order and count");
  }

  HostAcceptanceVerification verifyMinimumOnly() {
    return HostAcceptanceVerification.minimumOnly();
  }

  private HostAcceptanceSnapshot captureWorkbookRequirements(
      AcceptancePolicy.RequireAll authoredPolicy,
      AcceptancePolicy.RequireAll requireAll,
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation,
      @Nullable Path materializedSourceArtifact)
      throws IOException {
    List<HostAcceptanceSnapshot.SemanticPreservation> baselines = new ArrayList<>();
    List<HostAcceptanceSnapshot.OpaquePartPreservation> opaquePartBaselines = new ArrayList<>();
    for (int requirementIndex = 0;
        requirementIndex < requireAll.requirements().size();
        requirementIndex++) {
      var requirement = requireAll.requirements().get(requirementIndex);
      var original = originalRequirement(authoredPolicy, requireAll, requirementIndex);
      if (requirement instanceof AcceptancePolicy.Requirement.PreserveInspectionFacts facts) {
        baselines.add(
            new HostAcceptanceSnapshot.SemanticPreservation(
                facts,
                captureWorkbookInspections(
                    ((AcceptancePolicy.Requirement.PreserveInspectionFacts) original).inspections(),
                    facts.inspections(),
                    workbook,
                    workbookLocation)));
      }
      if (requirement instanceof AcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts parts) {
        if (materializedSourceArtifact == null) {
          throw new IOException(
              "host opaque-part preservation requires a materialized source workbook");
        }
        opaquePartBaselines.add(
            new HostAcceptanceSnapshot.OpaquePartPreservation(parts, materializedSourceArtifact));
      }
    }
    return new HostAcceptanceSnapshot(baselines, opaquePartBaselines);
  }

  private static List<WorkbookExecutionEvidence.Preservation> verifyStagedOpaquePartPreservation(
      AcceptancePolicy.RequireAll requireAll,
      HostAcceptanceSnapshot snapshot,
      Path stagedWorkbookArtifact,
      List<WorkbookExecutionEvidence.Preservation> preservation)
      throws IOException, PreservationFailedException {
    for (AcceptancePolicy.Requirement requirement : requireAll.requirements()) {
      if (requirement instanceof AcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts parts) {
        preservation.add(
            OpaqueOoxmlPartPreservation.establish(
                snapshot.sourceFor(parts), stagedWorkbookArtifact, parts.partNames()));
      }
    }
    return List.copyOf(preservation);
  }

  private void verifyWorkbookSemanticPreservation(
      AcceptancePolicy.Requirement.PreserveInspectionFacts original,
      AcceptancePolicy.Requirement.PreserveInspectionFacts facts,
      HostAcceptanceSnapshot snapshot,
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation,
      List<WorkbookExecutionEvidence.Preservation> preservation)
      throws IOException, PreservationFailedException {
    List<InspectionResult> observed =
        captureWorkbookInspections(
            original.inspections(), facts.inspections(), workbook, workbookLocation);
    establishSemanticPreservation(facts, snapshot.baselineFor(facts), observed, preservation);
  }

  private static void establishSemanticPreservation(
      AcceptancePolicy.Requirement.PreserveInspectionFacts facts,
      List<InspectionResult> baseline,
      List<InspectionResult> observed,
      List<WorkbookExecutionEvidence.Preservation> preservation)
      throws PreservationFailedException {
    List<dev.erst.gridgrind.core.vocabulary.StepId> stepIds =
        facts.inspections().stream().map(inspection -> inspection.stepId()).toList();
    if (!baseline.equals(observed)) {
      throw new PreservationFailedException(
          WorkbookExecutionEvidence.Preservation.Kind.SEMANTIC,
          stepIds.stream().map(dev.erst.gridgrind.core.vocabulary.StepId::value).toList());
    }
    preservation.add(new WorkbookExecutionEvidence.Preservation.SemanticEstablished(stepIds));
  }

  private List<InspectionResult> captureWorkbookInspections(
      List<dev.erst.gridgrind.core.plan.InspectionStep> originals,
      List<dev.erst.gridgrind.core.plan.InspectionStep> inspections,
      XssfWorkbook workbook,
      WorkbookLocation workbookLocation)
      throws IOException {
    List<InspectionResult> results = new ArrayList<>(inspections.size());
    for (int index = 0; index < inspections.size(); index++) {
      var inspection = inspections.get(index);
      results.add(
          stepSupport.executeInspectionStep(
              (dev.erst.gridgrind.core.plan.PreparedInspectionStep)
                  dev.erst.gridgrind.core.plan.PreparedStep.prepare(
                      originals.get(index), inspection),
              workbook,
              workbookLocation,
              new ExecutionModeSpec.FullXssf()));
    }
    return List.copyOf(results);
  }
}
