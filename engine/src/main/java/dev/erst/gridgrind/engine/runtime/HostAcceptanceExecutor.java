package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.WorkbookExecutionEvidence;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookLocation;
import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Evaluates host-owned terminal acceptance after mutations and before final publication. */
final class HostAcceptanceExecutor {
  private final ExecutionStepSupport stepSupport;

  HostAcceptanceExecutor(ExecutionStepSupport stepSupport) {
    this.stepSupport = Objects.requireNonNull(stepSupport, "stepSupport must not be null");
  }

  HostAcceptanceSnapshot captureBeforeMutation(
      GridGrindHostAcceptancePolicy policy,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      @Nullable Path materializedSourceArtifact)
      throws IOException {
    Objects.requireNonNull(workbook, "workbook must not be null");
    return switch (Objects.requireNonNull(policy, "policy must not be null")) {
      case GridGrindHostAcceptancePolicy.MinimumOnly _ -> HostAcceptanceSnapshot.empty();
      case GridGrindHostAcceptancePolicy.RequireAll requireAll ->
          captureWorkbookRequirements(
              requireAll, workbook, workbookLocation, materializedSourceArtifact);
    };
  }

  HostAcceptanceVerification verifyWorkbook(
      GridGrindHostAcceptancePolicy policy,
      HostAcceptanceSnapshot snapshot,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation)
      throws IOException, AssertionFailedException, PreservationFailedException {
    return switch (Objects.requireNonNull(policy, "policy must not be null")) {
      case GridGrindHostAcceptancePolicy.MinimumOnly _ -> verifyMinimumOnly();
      case GridGrindHostAcceptancePolicy.RequireAll requireAll ->
          verifyWorkbook(requireAll, snapshot, workbook, workbookLocation);
    };
  }

  /** Establishes any host-required opaque-part byte identity against a verified staged package. */
  List<WorkbookExecutionEvidence.Preservation> verifyStagedArtifact(
      GridGrindHostAcceptancePolicy policy,
      HostAcceptanceSnapshot snapshot,
      Path stagedWorkbookArtifact)
      throws IOException, PreservationFailedException {
    Objects.requireNonNull(snapshot, "snapshot must not be null");
    Objects.requireNonNull(stagedWorkbookArtifact, "stagedWorkbookArtifact must not be null");
    return switch (Objects.requireNonNull(policy, "policy must not be null")) {
      case GridGrindHostAcceptancePolicy.MinimumOnly _ -> List.of();
      case GridGrindHostAcceptancePolicy.RequireAll requireAll ->
          verifyStagedOpaquePartPreservation(requireAll, snapshot, stagedWorkbookArtifact);
    };
  }

  HostAcceptanceVerification verifyWorkbook(
      GridGrindHostAcceptancePolicy.RequireAll requireAll,
      HostAcceptanceSnapshot snapshot,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation)
      throws IOException, AssertionFailedException, PreservationFailedException {
    List<AssertionResult> assertions = new ArrayList<>();
    List<WorkbookExecutionEvidence.Preservation> preservation = new ArrayList<>();
    for (GridGrindHostAcceptancePolicy.Requirement requirement : requireAll.requirements()) {
      switch (requirement) {
        case GridGrindHostAcceptancePolicy.Requirement.TerminalAssertions terminalAssertions -> {
          for (var assertion : terminalAssertions.assertions()) {
            assertions.add(
                stepSupport.executeAssertionStep(
                    assertion,
                    workbook,
                    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null"),
                    new ExecutionModeInput.FullXssf()));
          }
        }
        case GridGrindHostAcceptancePolicy.Requirement.PreserveInspectionFacts facts ->
            verifyWorkbookSemanticPreservation(
                facts, snapshot, workbook, workbookLocation, preservation);
        case GridGrindHostAcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts _ -> {}
        case GridGrindHostAcceptancePolicy.Requirement.RequireCalculation _ -> {}
      }
    }
    return new HostAcceptanceVerification(assertions, preservation);
  }

  /**
   * Verifies the terminal-only acceptance subset available after mode-policy validation for an
   * event-read or streaming-write execution.
   */
  HostAcceptanceVerification verifyMaterializedTerminalAcceptance(
      GridGrindHostAcceptancePolicy policy, Path workbookPath, WorkbookLocation workbookLocation)
      throws IOException, AssertionFailedException {
    List<AssertionResult> assertions = new ArrayList<>();
    if (policy instanceof GridGrindHostAcceptancePolicy.RequireAll requireAll) {
      for (GridGrindHostAcceptancePolicy.Requirement requirement : requireAll.requirements()) {
        if (requirement
            instanceof GridGrindHostAcceptancePolicy.Requirement.TerminalAssertions terminal) {
          for (var assertion : terminal.assertions()) {
            assertions.add(
                stepSupport.executeAssertionAgainstMaterializedPath(
                    assertion,
                    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null"),
                    workbookPath));
          }
        }
      }
    }
    return new HostAcceptanceVerification(assertions, List.of());
  }

  HostAcceptanceVerification verifyMinimumOnly() {
    return HostAcceptanceVerification.minimumOnly();
  }

  private HostAcceptanceSnapshot captureWorkbookRequirements(
      GridGrindHostAcceptancePolicy.RequireAll requireAll,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      @Nullable Path materializedSourceArtifact)
      throws IOException {
    List<HostAcceptanceSnapshot.SemanticPreservation> baselines = new ArrayList<>();
    List<HostAcceptanceSnapshot.OpaquePartPreservation> opaquePartBaselines = new ArrayList<>();
    for (GridGrindHostAcceptancePolicy.Requirement requirement : requireAll.requirements()) {
      if (requirement
          instanceof GridGrindHostAcceptancePolicy.Requirement.PreserveInspectionFacts facts) {
        baselines.add(
            new HostAcceptanceSnapshot.SemanticPreservation(
                facts,
                captureWorkbookInspections(facts.inspections(), workbook, workbookLocation)));
      }
      if (requirement
          instanceof GridGrindHostAcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts parts) {
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
      GridGrindHostAcceptancePolicy.RequireAll requireAll,
      HostAcceptanceSnapshot snapshot,
      Path stagedWorkbookArtifact)
      throws IOException, PreservationFailedException {
    List<WorkbookExecutionEvidence.Preservation> preservation = new ArrayList<>();
    for (GridGrindHostAcceptancePolicy.Requirement requirement : requireAll.requirements()) {
      if (requirement
          instanceof GridGrindHostAcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts parts) {
        preservation.add(
            OpaqueOoxmlPartPreservation.establish(
                snapshot.sourceFor(parts), stagedWorkbookArtifact, parts.partNames()));
      }
    }
    return List.copyOf(preservation);
  }

  private void verifyWorkbookSemanticPreservation(
      GridGrindHostAcceptancePolicy.Requirement.PreserveInspectionFacts facts,
      HostAcceptanceSnapshot snapshot,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      List<WorkbookExecutionEvidence.Preservation> preservation)
      throws IOException, PreservationFailedException {
    List<InspectionResult> observed =
        captureWorkbookInspections(facts.inspections(), workbook, workbookLocation);
    establishSemanticPreservation(facts, snapshot.baselineFor(facts), observed, preservation);
  }

  private static void establishSemanticPreservation(
      GridGrindHostAcceptancePolicy.Requirement.PreserveInspectionFacts facts,
      List<InspectionResult> baseline,
      List<InspectionResult> observed,
      List<WorkbookExecutionEvidence.Preservation> preservation)
      throws PreservationFailedException {
    List<String> stepIds =
        facts.inspections().stream().map(inspection -> inspection.stepId()).toList();
    if (!baseline.equals(observed)) {
      throw new PreservationFailedException(
          WorkbookExecutionEvidence.Preservation.Kind.SEMANTIC, stepIds);
    }
    preservation.add(new WorkbookExecutionEvidence.Preservation.SemanticEstablished(stepIds));
  }

  private List<InspectionResult> captureWorkbookInspections(
      List<dev.erst.gridgrind.contract.step.InspectionStep> inspections,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation)
      throws IOException {
    List<InspectionResult> results = new ArrayList<>(inspections.size());
    for (var inspection : inspections) {
      results.add(
          stepSupport.executeInspectionStep(
              inspection, workbook, workbookLocation, new ExecutionModeInput.FullXssf()));
    }
    return List.copyOf(results);
  }
}
