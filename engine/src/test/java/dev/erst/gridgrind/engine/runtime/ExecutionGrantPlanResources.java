package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import java.nio.file.Path;
import java.util.List;

/** Derives the host resources and publication authority required by one generated plan. */
final class ExecutionGrantPlanResources {
  private ExecutionGrantPlanResources() {}

  static List<Path> readableFiles(WorkbookPlan plan, Path workingDirectory) {
    java.util.stream.Stream<Path> source =
        switch (plan.source()) {
          case WorkbookPlan.WorkbookSource.New _ -> java.util.stream.Stream.empty();
          case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
              java.util.stream.Stream.of(resolve(existingFile.path(), workingDirectory));
        };
    java.util.stream.Stream<Path> externalWorkbooks =
        plan.formulaEnvironment().externalWorkbooks().stream()
            .map(workbook -> resolve(workbook.path(), workingDirectory));
    return java.util.stream.Stream.concat(source, externalWorkbooks).distinct().toList();
  }

  static GridGrindExecutionGrant.PublicationAuthority publicationAuthority(
      WorkbookPlan.WorkbookPersistence persistence, Path workingDirectory) {
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ ->
          new GridGrindExecutionGrant.PublicationAuthority.None();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          new GridGrindExecutionGrant.PublicationAuthority.SaveAs(
              resolve(saveAs.path(), workingDirectory), saveAs.ifExists());
      case WorkbookPlan.WorkbookPersistence.Overwrite _ ->
          new GridGrindExecutionGrant.PublicationAuthority.OverwriteSource();
    };
  }

  private static Path resolve(String rawPath, Path workingDirectory) {
    Path candidate = Path.of(rawPath);
    return (candidate.isAbsolute() ? candidate : workingDirectory.resolve(candidate))
        .toAbsolutePath()
        .normalize();
  }
}
