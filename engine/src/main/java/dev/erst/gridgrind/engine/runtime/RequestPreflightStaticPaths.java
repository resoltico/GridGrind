package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.util.List;
import java.util.Objects;

/**
 * Validates request-owned path containment without binding, opening, or creating filesystem state.
 */
final class RequestPreflightStaticPaths {
  private RequestPreflightStaticPaths() {}

  static void validate(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(bindings, "bindings must not be null");
    Objects.requireNonNull(problems, "problems must not be null");
    validateSourcePath(request, bindings, problems);
    validatePersistencePath(request, bindings, problems);
    validateFormulaWorkbookPaths(request, bindings, problems);
  }

  private static void validateSourcePath(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    if (request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile) {
      validate(
          existingFile.path(),
          bindings,
          new ProblemContext.OpenWorkbook(
              ExecutionRequestPaths.requestShape(request),
              ExecutionRequestPaths.workbookReference(request, bindings.workingDirectory())),
          problems);
    }
  }

  private static void validatePersistencePath(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    if (request.persistence() instanceof WorkbookPlan.WorkbookPersistence.SaveAs saveAs) {
      validate(
          saveAs.path(),
          bindings,
          new ProblemContext.PersistWorkbook(
              ExecutionRequestPaths.requestShape(request),
              new dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces
                  .PersistenceReference.SaveAs(saveAs.path())),
          problems);
    }
  }

  private static void validateFormulaWorkbookPaths(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    for (var externalWorkbook : request.formulaEnvironment().externalWorkbooks()) {
      validateFormulaWorkbookPath(request, externalWorkbook.path(), bindings, problems);
    }
  }

  private static void validateFormulaWorkbookPath(
      WorkbookPlan request,
      String path,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    validate(
        path,
        bindings,
        new ProblemContext.ResolveInputs(
            ExecutionRequestPaths.requestShape(request),
            dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.InputReference.path(
                "formula external workbook", path)),
        problems);
  }

  private static void validate(
      String path,
      ExecutionInputBindings bindings,
      ProblemContext context,
      List<GridGrindProblemDetail.Problem> problems) {
    try {
      ExecutionRequestPaths.normalizePath(path, bindings.workingDirectory());
    } catch (Exception exception) {
      problems.add(GridGrindProblems.fromException(exception, context));
    }
  }
}
