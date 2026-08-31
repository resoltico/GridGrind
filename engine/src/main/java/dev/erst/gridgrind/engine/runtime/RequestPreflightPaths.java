package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.InputReference;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceSignature;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlSigningMaterialSupport;
import java.io.IOException;
import java.util.List;

/** Preflights request-owned source, signing, and persistence paths without writing a workbook. */
final class RequestPreflightPaths {
  private RequestPreflightPaths() {}

  static RequestPreflightPaths newForVerification() {
    return new RequestPreflightPaths();
  }

  static void verify(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    preflightFormulaEnvironmentPaths(request, bindings, problems);
    preflightPersistenceMaterial(request, bindings, problems);
    preflightPersistenceTarget(request, bindings, problems);
  }

  private static void preflightFormulaEnvironmentPaths(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    for (var externalWorkbook : request.formulaEnvironment().externalWorkbooks()) {
      try {
        bindings
            .requestPathAccess()
            .materializeRead(
                externalWorkbook.path(),
                "formulaEnvironment.externalWorkbooks",
                "gridgrind-formula-workbook-",
                ".xlsx");
      } catch (IOException exception) {
        addFormulaPathProblem(
            request,
            externalWorkbook,
            SourceBackedPlanResolver.inputFileFailure(
                externalWorkbook.path(), "formula external workbook", exception),
            problems);
      } catch (RuntimeException exception) {
        addFormulaPathProblem(request, externalWorkbook, exception, problems);
      }
    }
  }

  private static void addFormulaPathProblem(
      WorkbookPlan request,
      dev.erst.gridgrind.contract.dto.FormulaExternalWorkbookInput externalWorkbook,
      Exception exception,
      List<GridGrindProblemDetail.Problem> problems) {
    problems.add(
        GridGrindProblems.fromException(
            exception,
            new ProblemContext.ResolveInputs(
                ExecutionRequestPaths.requestShape(request),
                InputReference.path("formula external workbook", externalWorkbook.path()))));
  }

  static void preflightPersistenceMaterial(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    try {
      ExcelOoxmlPersistenceOptions options =
          ExecutionRequestPaths.persistenceOptions(request.persistence(), bindings);
      verifySigningMaterial(options.signature());
    } catch (Exception exception) {
      problems.add(
          GridGrindProblems.fromException(
              exception,
              new ProblemContext.PersistWorkbook(
                  ExecutionRequestPaths.requestShape(request),
                  ExecutionRequestPaths.persistenceReference(
                      request, bindings.workingDirectory()))));
    }
  }

  static void verifySigningMaterial(ExcelOoxmlPersistenceSignature signature) {
    if (signature instanceof ExcelOoxmlPersistenceSignature.Sign sign) {
      ExcelOoxmlSigningMaterialSupport.signingMaterial(sign.options());
    }
  }

  private static void preflightPersistenceTarget(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    java.util.Optional<PersistenceTarget> target = persistenceTarget(request);
    if (target.isEmpty()) {
      return;
    }
    try {
      PersistenceTarget resolvedTarget = target.orElseThrow();
      bindings
          .requestPathAccess()
          .prepareOutput(resolvedTarget.path(), "persistence", resolvedTarget.disposition());
    } catch (Exception exception) {
      problems.add(
          GridGrindProblems.fromException(
              exception,
              new ProblemContext.PersistWorkbook(
                  ExecutionRequestPaths.requestShape(request),
                  ExecutionRequestPaths.persistenceReference(
                      request, bindings.workingDirectory()))));
    }
  }

  static java.util.Optional<PersistenceTarget> persistenceTarget(WorkbookPlan request) {
    return switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.None _ -> java.util.Optional.empty();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          java.util.Optional.of(
              new PersistenceTarget(saveAs.path(), ExecutionRequestPaths.writeDisposition(saveAs)));
      case WorkbookPlan.WorkbookPersistence.Overwrite _ ->
          request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile
              ? java.util.Optional.of(
                  new PersistenceTarget(
                      existingFile.path(), WorkbookArtifactWriteDisposition.REPLACE_EXISTING))
              : java.util.Optional.empty();
    };
  }

  record PersistenceTarget(String path, WorkbookArtifactWriteDisposition disposition) {}
}
