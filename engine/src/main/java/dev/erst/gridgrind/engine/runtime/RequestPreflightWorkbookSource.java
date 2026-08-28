package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.step.WorkbookStaticRequestContract;
import dev.erst.gridgrind.contract.step.WorkbookStaticViolation;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPackagePersistenceSupport;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceEncryption;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceOptions;
import dev.erst.gridgrind.excel.ooxml.ExcelOoxmlPersistenceSignature;
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
    ExecutionWorkbookSupport workbookSupport =
        new ExecutionWorkbookSupport(bindings.tempFileFactory());
    try (ExcelWorkbook workbook = workbookSupport.openWorkbook(request.source(), null, bindings)) {
      validateSourcePersistencePolicy(request, workbook);
      addFormulaColumnViolations(request, bindings, workbook, problems);
    } catch (Exception exception) {
      problems.add(GridGrindProblems.fromException(exception, context));
    }
  }

  private static void addFormulaColumnViolations(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      ExcelWorkbook workbook,
      List<GridGrindProblemDetail.Problem> problems) {
    for (WorkbookStaticViolation violation :
        WorkbookStaticRequestContract.validateKnownFormulaPresence(
            request, containsFormula(workbook))) {
      problems.add(
          GridGrindProblems.problem(
              dev.erst.gridgrind.contract.dto.GridGrindProblemCode.INVALID_REQUEST,
              violation.message(),
              RequestPreflight.openWorkbookContext(request, bindings),
              List.of()));
    }
  }

  private static boolean containsFormula(ExcelWorkbook workbook) {
    for (org.apache.poi.ss.usermodel.Sheet sheet : workbook.xssfWorkbook()) {
      for (org.apache.poi.ss.usermodel.Row row : sheet) {
        for (org.apache.poi.ss.usermodel.Cell cell : row) {
          if (cell.getCellType() == org.apache.poi.ss.usermodel.CellType.FORMULA) {
            return true;
          }
        }
      }
    }
    return false;
  }

  private static void validateSourcePersistencePolicy(
      WorkbookPlan request, ExcelWorkbook workbook) {
    java.util.Optional<dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput> security =
        switch (request.persistence()) {
          case WorkbookPlan.WorkbookPersistence.None _ -> java.util.Optional.empty();
          case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> saveAs.security();
          case WorkbookPlan.WorkbookPersistence.Overwrite overwrite -> overwrite.security();
        };
    if (security.isPresent()
        && security.orElseThrow().encryption()
            instanceof
            dev.erst.gridgrind.contract.dto.OoxmlPersistenceEncryptionInput.PreserveSource
        && workbook.persistence().loadedPackageSecurity().encryption()
            instanceof dev.erst.gridgrind.excel.ooxml.ExcelOoxmlEncryptionSnapshot.None) {
      throw new EncryptionSourceNotEncryptedException();
    }
    if (security.isPresent()
        && security.orElseThrow().encryption()
            instanceof
            dev.erst.gridgrind.contract.dto.OoxmlPersistenceEncryptionInput.PreserveSource) {
      ExcelOoxmlPackagePersistenceSupport.effectiveOptions(
          workbook.persistence().loadedPackageSecurity(),
          workbook.persistence().sourceEncryptionPassword(),
          new ExcelOoxmlPersistenceOptions(
              new ExcelOoxmlPersistenceEncryption.PreserveSource(),
              new ExcelOoxmlPersistenceSignature.Unsigned()));
    }
  }
}
