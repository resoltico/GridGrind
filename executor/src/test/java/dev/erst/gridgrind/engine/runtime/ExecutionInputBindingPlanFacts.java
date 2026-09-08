package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
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
      dev.erst.gridgrind.contract.dto.WorkbookPlan.WorkbookPersistence persistence,
      Path workingDirectory) {
    return switch (persistence) {
      case dev.erst.gridgrind.contract.dto.WorkbookPlan.WorkbookPersistence.None _ ->
          new GridGrindExecutionGrant.PublicationAuthority.None();
      case dev.erst.gridgrind.contract.dto.WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          new GridGrindExecutionGrant.PublicationAuthority.SaveAs(
              resolve(saveAs.path(), workingDirectory), saveAs.ifExists());
      case dev.erst.gridgrind.contract.dto.WorkbookPlan.WorkbookPersistence.Overwrite _ ->
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
