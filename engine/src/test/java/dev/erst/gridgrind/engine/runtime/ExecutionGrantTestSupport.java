package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import java.nio.file.Path;
import java.util.List;

/** Builds explicit, narrowly scoped host grants for engine-runtime tests. */
final class ExecutionGrantTestSupport {
  private ExecutionGrantTestSupport() {}

  static GridGrindExecutionGrant.Bounded noPublication() {
    return grant(List.of(), List.of());
  }

  static GridGrindExecutionGrant.Bounded noPublication(
      List<String> operationIds, List<Path> readableFiles) {
    return grant(
        operationIds,
        readableFiles.stream()
            .<GridGrindExecutionGrant.ReadAuthority>map(
                GridGrindExecutionGrant.ReadAuthority.File::new)
            .toList());
  }

  static GridGrindExecutionGrant.Bounded noPublicationWithStandardInput(List<String> operationIds) {
    return grant(operationIds, List.of(new GridGrindExecutionGrant.ReadAuthority.StandardInput()));
  }

  static GridGrindExecutionGrant.Bounded forPlan(
      WorkbookPlan plan, Path workingDirectory, Path... additionalReadableFiles) {
    Path normalizedWorkingDirectory = workingDirectory.toAbsolutePath().normalize();
    List<Path> readableFiles =
        java.util.stream.Stream.concat(
                ExecutionGrantPlanResources.readableFiles(plan, normalizedWorkingDirectory)
                    .stream(),
                java.util.Arrays.stream(additionalReadableFiles)
                    .map(path -> path.toAbsolutePath().normalize()))
            .toList();
    List<GridGrindExecutionGrant.ReadAuthority> readableResources =
        readableFiles.stream()
            .<GridGrindExecutionGrant.ReadAuthority>map(
                GridGrindExecutionGrant.ReadAuthority.File::new)
            .toList();
    return new GridGrindExecutionGrant.Bounded(
        readableResources,
        plan.steps().stream().map(ExecutionGrantPlanOperationIds::forStep).distinct().toList(),
        new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        ExecutionGrantPlanResources.publicationAuthority(
            plan.persistence(), normalizedWorkingDirectory),
        ExecutionGrantPlanSecrets.requestedBy(plan),
        dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());
  }

  static GridGrindExecutionGrant.Bounded overwriteSource(
      List<String> operationIds, List<Path> readableFiles) {
    return new GridGrindExecutionGrant.Bounded(
        readableFiles.stream()
            .<GridGrindExecutionGrant.ReadAuthority>map(
                GridGrindExecutionGrant.ReadAuthority.File::new)
            .toList(),
        operationIds,
        new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        new GridGrindExecutionGrant.PublicationAuthority.OverwriteSource(),
        List.of(),
        dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());
  }

  static GridGrindExecutionGrant.Bounded saveAs(
      List<String> operationIds,
      List<Path> readableFiles,
      Path output,
      WorkbookPlan.WorkbookPersistence.IfExists ifExists) {
    return new GridGrindExecutionGrant.Bounded(
        readableFiles.stream()
            .<GridGrindExecutionGrant.ReadAuthority>map(
                GridGrindExecutionGrant.ReadAuthority.File::new)
            .toList(),
        operationIds,
        new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        new GridGrindExecutionGrant.PublicationAuthority.SaveAs(output, ifExists),
        List.of(),
        dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());
  }

  private static GridGrindExecutionGrant.Bounded grant(
      List<String> operationIds, List<GridGrindExecutionGrant.ReadAuthority> readableResources) {
    return new GridGrindExecutionGrant.Bounded(
        readableResources,
        operationIds,
        new GridGrindExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        new GridGrindExecutionGrant.PublicationAuthority.None(),
        List.of(),
        dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy.minimum());
  }
}
