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

/** Prepares the selected source path and validates facts required by that concrete strategy. */
final class RequestPreflightWorkbookSource {
  private RequestPreflightWorkbookSource() {}

  static PreparedWorkbookSource prepare(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    return prepare(
        request,
        bindings,
        problems,
        (plan, inputs) ->
            new WorkbookLifecycle(inputs.tempFileFactory())
                .openWorkbook(
                    plan.source(), problems.isEmpty() ? plan.formulaEnvironment() : null, inputs));
  }

  @SuppressWarnings(
      "PMD.CloseResource") // Both native handles transfer immediately to the session-owned prepared
  // source.
  static PreparedWorkbookSource prepare(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems,
      XssfOpener opener) {
    var prepared =
        new PreparedWorkbookSource(
            dev.erst.gridgrind.core.semantics.BackendCapabilities.executionStrategy(
                request, bindings.executionGrant()));
    bindings.own(prepared::close);
    if (!(request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile existing)) {
      prepared.newWorkbookCalendar();
      return prepared;
    }
    var context = RequestPreflight.openWorkbookContext(request, bindings);
    try {
      if (prepared.strategy().backend()
          == dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Backend.EVENT_READ) {
        var materialized =
            dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactIo.materializeWorkbook(
                WorkbookLifecycle.materializeSource(existing, bindings),
                ExecutionRequestPaths.materializeOpenSecurity(
                    existing.security().orElse(null), bindings),
                bindings.tempFileFactory()::createTempFile);
        prepared.retain(materialized);
        // EVENT_READ admits inspection steps only. The formula/column-edit rule has no
        // applicable mutation here; security and workbook metadata still require validation.
        new dev.erst.gridgrind.engine.backend.eventread.EventWorkbookReader()
            .apply(materialized.workbookPath(), List.of());
      } else {
        XssfWorkbook workbook = opener.open(request, bindings);
        prepared.retain(workbook);
        validateSourcePersistencePolicy(request, workbook);
        addFormulaColumnViolations(request, bindings, workbook, problems);
      }
    } catch (Exception exception) {
      problems.add(GridGrindProblems.fromException(exception, context));
    }
    return prepared;
  }

  /** Package-local native opening seam for source-lifetime and strategy regression controls. */
  @FunctionalInterface
  interface XssfOpener {
    /** Opens a source whose exclusive close obligation transfers to the prepared-source owner. */
    XssfWorkbook open(WorkbookPlan request, ExecutionInputBindings bindings)
        throws java.io.IOException;
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
