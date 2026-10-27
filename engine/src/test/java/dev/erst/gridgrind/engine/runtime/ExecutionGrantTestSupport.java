package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import java.nio.file.Path;
import java.util.List;

/** Builds explicit, narrowly scoped host grants for engine-runtime tests. */
final class ExecutionGrantTestSupport {
  private ExecutionGrantTestSupport() {}

  static ExecutionGrant noPublication() {
    return grant(List.of(), List.of());
  }

  static ExecutionGrant noPublication(List<String> operationIds, List<Path> readableFiles) {
    return grant(
        operationIds,
        readableFiles.stream()
            .<ExecutionGrant.ReadAuthority>map(ExecutionGrant.ReadAuthority.File::new)
            .toList());
  }

  static ExecutionGrant noPublicationWithStandardInput(List<String> operationIds) {
    return grant(operationIds, List.of(new ExecutionGrant.ReadAuthority.Stdin()));
  }

  static ExecutionGrant forPlan(
      WorkbookPlan plan, Path workingDirectory, Path... additionalReadableFiles) {
    Path normalizedWorkingDirectory = workingDirectory.toAbsolutePath().normalize();
    List<Path> readableFiles =
        java.util.stream.Stream.concat(
                ExecutionGrantPlanResources.readableFiles(plan, normalizedWorkingDirectory)
                    .stream(),
                java.util.Arrays.stream(additionalReadableFiles)
                    .map(path -> path.toAbsolutePath().normalize()))
            .toList();
    List<ExecutionGrant.ReadAuthority> readableResources =
        readableFiles.stream()
            .<ExecutionGrant.ReadAuthority>map(ExecutionGrant.ReadAuthority.File::new)
            .toList();
    return new ExecutionGrant(
        readableResources,
        plan.steps().stream().map(ExecutionGrantPlanOperationIds::forStep).distinct().toList(),
        new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        ExecutionGrantPlanResources.publicationAuthority(
            plan.persistence(), normalizedWorkingDirectory),
        ExecutionGrantPlanSecrets.requestedBy(plan),
        dev.erst.gridgrind.core.grant.AcceptancePolicy.minimum());
  }

  static ExecutionGrant overwriteSource(List<String> operationIds, List<Path> readableFiles) {
    return new ExecutionGrant(
        readableFiles.stream()
            .<ExecutionGrant.ReadAuthority>map(ExecutionGrant.ReadAuthority.File::new)
            .toList(),
        operationIds,
        new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        new ExecutionGrant.PublicationAuthority.OverwriteSource(),
        List.of(),
        dev.erst.gridgrind.core.grant.AcceptancePolicy.minimum());
  }

  static ExecutionGrant saveAs(
      List<String> operationIds,
      List<Path> readableFiles,
      Path output,
      WorkbookPlan.WorkbookPersistence.IfExists ifExists) {
    return new ExecutionGrant(
        readableFiles.stream()
            .<ExecutionGrant.ReadAuthority>map(ExecutionGrant.ReadAuthority.File::new)
            .toList(),
        operationIds,
        new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        new ExecutionGrant.PublicationAuthority.SaveAs(output, ifExists),
        List.of(),
        dev.erst.gridgrind.core.grant.AcceptancePolicy.minimum());
  }

  private static ExecutionGrant grant(
      List<String> operationIds, List<ExecutionGrant.ReadAuthority> readableResources) {
    return new ExecutionGrant(
        readableResources,
        operationIds,
        new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
        new ExecutionGrant.PublicationAuthority.None(),
        List.of(),
        dev.erst.gridgrind.core.grant.AcceptancePolicy.minimum());
  }
}
