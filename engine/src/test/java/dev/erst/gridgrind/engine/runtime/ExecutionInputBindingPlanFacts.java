package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.MutationStep;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import java.nio.file.Path;

/** Derives operation and publication facts for explicit executor-test input bindings. */
final class ExecutionInputBindingPlanFacts {
  private ExecutionInputBindingPlanFacts() {}

  static String operationId(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutation -> mutation.action().actionType();
      case AssertionStep assertion -> assertion.assertion().assertionType();
      case InspectionStep inspection -> inspection.query().queryType();
    };
  }

  static GridGrindExecutionGrant.PublicationAuthority publicationAuthority(
      dev.erst.gridgrind.core.plan.WorkbookPlan.WorkbookPersistence persistence,
      Path workingDirectory) {
    return switch (persistence) {
      case dev.erst.gridgrind.core.plan.WorkbookPlan.WorkbookPersistence.None _ ->
          new GridGrindExecutionGrant.PublicationAuthority.None();
      case dev.erst.gridgrind.core.plan.WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          new GridGrindExecutionGrant.PublicationAuthority.SaveAs(
              resolve(saveAs.path(), workingDirectory), saveAs.ifExists());
      case dev.erst.gridgrind.core.plan.WorkbookPlan.WorkbookPersistence.Overwrite _ ->
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
