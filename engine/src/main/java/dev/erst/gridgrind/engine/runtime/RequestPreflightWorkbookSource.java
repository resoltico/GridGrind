package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec;
import dev.erst.gridgrind.core.model.security.OoxmlPersistenceSignatureSpec;
import dev.erst.gridgrind.core.model.security.ResolvedPersistenceSecurity;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.validation.WorkbookStaticRequestContract;
import dev.erst.gridgrind.core.validation.WorkbookStaticViolation;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.security.SecurityAdapter;
import java.util.List;

/**
 * Opens an existing source only long enough to validate non-mutating source-dependent constraints.
 */
final class RequestPreflightWorkbookSource {
  private RequestPreflightWorkbookSource() {}

  static void verify(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    if (!(request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile)) {
      return;
    }
    var context = RequestPreflight.openWorkbookContext(request, bindings);
    WorkbookLifecycle workbookSupport = new WorkbookLifecycle(bindings.tempFileFactory());
    try (XssfWorkbook workbook = workbookSupport.openWorkbook(request.source(), null, bindings)) {
      validateSourcePersistencePolicy(request, workbook);
      addFormulaColumnViolations(request, bindings, workbook, problems);
    } catch (Exception exception) {
      problems.add(GridGrindProblems.fromException(exception, context));
    }
  }

  private static void addFormulaColumnViolations(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      XssfWorkbook workbook,
      List<GridGrindProblemDetail.Problem> problems) {
    for (WorkbookStaticViolation violation :
        WorkbookStaticRequestContract.validateKnownFormulaPresence(
            request, workbook.formulas().containsFormula())) {
      problems.add(
          GridGrindProblems.problem(
              dev.erst.gridgrind.core.problem.GridGrindProblemCode.INVALID_REQUEST,
              violation.message(),
              RequestPreflight.openWorkbookContext(request, bindings),
              List.of()));
    }
  }

  private static void validateSourcePersistencePolicy(WorkbookPlan request, XssfWorkbook workbook) {
    java.util.Optional<dev.erst.gridgrind.core.model.security.OoxmlPersistenceSecuritySpec>
        security =
            switch (request.persistence()) {
              case WorkbookPlan.WorkbookPersistence.None _ -> java.util.Optional.empty();
              case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> saveAs.security();
              case WorkbookPlan.WorkbookPersistence.Overwrite overwrite -> overwrite.security();
            };
    if (security.isPresent()
        && security.orElseThrow().encryption()
            instanceof
            dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec.PreserveSource
        && workbook.persistence().loadedPackageSecurity().encryption()
            instanceof dev.erst.gridgrind.core.model.security.OoxmlEncryptionReport.None) {
      throw new EncryptionSourceNotEncryptedException();
    }
    if (security.isPresent()
        && security.orElseThrow().encryption()
            instanceof
            dev.erst.gridgrind.core.model.security.OoxmlPersistenceEncryptionSpec.PreserveSource) {
      SecurityAdapter.effectiveOptions(
          workbook.persistence().loadedPackageSecurity(),
          workbook.persistence().sourceEncryptionPassword(),
          new ResolvedPersistenceSecurity(
              new OoxmlPersistenceEncryptionSpec.PreserveSource(),
              new OoxmlPersistenceSignatureSpec.None()));
    }
  }
}
