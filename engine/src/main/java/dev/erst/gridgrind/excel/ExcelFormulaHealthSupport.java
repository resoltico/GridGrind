package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.AnalysisFindingCode;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;

/** Derived health analysis for authored formulas on one sheet. */
final class ExcelFormulaHealthSupport {
  private ExcelFormulaHealthSupport() {}

  static List<WorkbookAnalysis.AnalysisFinding> formulaHealthFindings(
      String sheetName,
      org.apache.poi.ss.usermodel.Sheet sheet,
      ExcelFormulaRuntime formulaRuntime) {
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    ExcelFormulaRuntimeContext formulaContext = formulaRuntime.context();
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (cell.getCellType() == CellType.FORMULA) {
          analyzeFormulaCell(findings, sheetName, cell, formulaRuntime, formulaContext);
        }
      }
    }
    return List.copyOf(findings);
  }

  private static void analyzeFormulaCell(
      List<WorkbookAnalysis.AnalysisFinding> findings,
      String sheetName,
      Cell cell,
      ExcelFormulaRuntime formulaRuntime,
      ExcelFormulaRuntimeContext formulaContext) {
    String address = cell.getAddress().formatAsString();
    String formula = cell.getCellFormula();
    WorkbookAnalysis.AnalysisLocation.Cell location =
        new WorkbookAnalysis.AnalysisLocation.Cell(sheetName, address);
    boolean missingExternalWorkbook =
        addExternalWorkbookFindings(findings, formulaContext, formula, location);
    addUserDefinedFunctionFinding(findings, formulaContext, formula, location);
    addVolatileFunctionFindings(findings, formula, location);
    addEvaluationFindings(
        findings, formulaRuntime, formulaContext, cell, formula, location, missingExternalWorkbook);
  }

  private static boolean addExternalWorkbookFindings(
      List<WorkbookAnalysis.AnalysisFinding> findings,
      ExcelFormulaRuntimeContext formulaContext,
      String formula,
      WorkbookAnalysis.AnalysisLocation.Cell location) {
    boolean missingExternalWorkbook = false;
    for (String workbookName : FormulaExceptions.externalWorkbookNames(formula)) {
      if (formulaContext.hasExternalWorkbookBinding(workbookName)) {
        continue;
      }
      if (formulaContext.missingWorkbookPolicy()
          == ExcelFormulaMissingWorkbookPolicy.USE_CACHED_VALUE) {
        findings.add(cachedExternalWorkbookFinding(formula, workbookName, location));
      } else {
        missingExternalWorkbook = true;
        findings.add(missingExternalWorkbookFinding(formula, workbookName, location));
      }
    }
    return missingExternalWorkbook;
  }

  private static WorkbookAnalysis.AnalysisFinding cachedExternalWorkbookFinding(
      String formula, String workbookName, WorkbookAnalysis.AnalysisLocation.Cell location) {
    return new WorkbookAnalysis.AnalysisFinding(
        AnalysisFindingCode.FORMULA_USES_CACHED_EXTERNAL_VALUE,
        AnalysisSeverity.INFO,
        "External workbook uses cached result",
        "Formula references external workbook "
            + workbookName
            + " and will fall back to the cached formula result when the workbook is missing.",
        location,
        List.of(formula, workbookName));
  }

  private static WorkbookAnalysis.AnalysisFinding missingExternalWorkbookFinding(
      String formula, String workbookName, WorkbookAnalysis.AnalysisLocation.Cell location) {
    return new WorkbookAnalysis.AnalysisFinding(
        AnalysisFindingCode.FORMULA_MISSING_EXTERNAL_WORKBOOK,
        AnalysisSeverity.ERROR,
        "External workbook is missing or unbound",
        "Formula references external workbook " + workbookName + " but no binding is configured.",
        location,
        List.of(formula, workbookName));
  }

  private static void addUserDefinedFunctionFinding(
      List<WorkbookAnalysis.AnalysisFinding> findings,
      ExcelFormulaRuntimeContext formulaContext,
      String formula,
      WorkbookAnalysis.AnalysisLocation.Cell location) {
    String leadingFunctionName = FormulaExceptions.leadingFunctionName(formula).orElse(null);
    if (!ExcelSheetAnalysisSupport.hasUnregisteredUserDefinedFunction(
        formulaContext, leadingFunctionName, formula)) {
      return;
    }
    findings.add(
        new WorkbookAnalysis.AnalysisFinding(
            AnalysisFindingCode.FORMULA_UNREGISTERED_USER_DEFINED_FUNCTION,
            AnalysisSeverity.ERROR,
            "User-defined function is not registered",
            "Formula references user-defined function "
                + leadingFunctionName
                + " but the current formula environment does not register it.",
            location,
            List.of(formula, leadingFunctionName)));
  }

  private static void addVolatileFunctionFindings(
      List<WorkbookAnalysis.AnalysisFinding> findings,
      String formula,
      WorkbookAnalysis.AnalysisLocation.Cell location) {
    ExcelSheetAnalysisSupport.volatileFunctions(formula)
        .forEach(
            functionName ->
                findings.add(
                    new WorkbookAnalysis.AnalysisFinding(
                        AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                        AnalysisSeverity.INFO,
                        "Volatile formula function",
                        "Formula uses volatile function " + functionName + ".",
                        location,
                        List.of(formula, functionName))));
  }

  private static void addEvaluationFindings(
      List<WorkbookAnalysis.AnalysisFinding> findings,
      ExcelFormulaRuntime formulaRuntime,
      ExcelFormulaRuntimeContext formulaContext,
      Cell cell,
      String formula,
      WorkbookAnalysis.AnalysisLocation.Cell location,
      boolean missingExternalWorkbook) {
    try {
      var evaluated = formulaRuntime.evaluate(cell);
      if (evaluated != null && evaluated.getCellType() == CellType.ERROR) {
        findings.add(
            new WorkbookAnalysis.AnalysisFinding(
                AnalysisFindingCode.FORMULA_ERROR_RESULT,
                AnalysisSeverity.ERROR,
                "Formula evaluates to an error",
                "Formula currently evaluates to "
                    + FormulaError.forInt(evaluated.getErrorValue()).getString()
                    + ".",
                location,
                List.of(formula)));
      }
    } catch (RuntimeException exception) {
      addEvaluationFailureFinding(
          findings, formulaContext, formula, location, missingExternalWorkbook, exception);
    }
  }

  private static void addEvaluationFailureFinding(
      List<WorkbookAnalysis.AnalysisFinding> findings,
      ExcelFormulaRuntimeContext formulaContext,
      String formula,
      WorkbookAnalysis.AnalysisLocation.Cell location,
      boolean missingExternalWorkbook,
      RuntimeException exception) {
    if (FormulaExceptions.isMissingExternalWorkbookFailure(exception)) {
      if (!missingExternalWorkbook) {
        String workbookName = FormulaExceptions.missingExternalWorkbookName(exception, formula);
        findings.add(
            new WorkbookAnalysis.AnalysisFinding(
                AnalysisFindingCode.FORMULA_MISSING_EXTERNAL_WORKBOOK,
                AnalysisSeverity.ERROR,
                "External workbook is missing or unbound",
                "Formula evaluation failed because external workbook "
                    + workbookName
                    + " could not be resolved.",
                location,
                List.of(formula, workbookName == null ? "" : workbookName)));
      }
      return;
    }
    if (!FormulaExceptions.isUnregisteredUserDefinedFunctionFailure(
        formulaContext, exception, formula)) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              AnalysisFindingCode.FORMULA_EVALUATION_FAILURE,
              AnalysisSeverity.ERROR,
              "Formula evaluation failed",
              "Formula evaluation failed: " + ExcelSheetAnalysisSupport.exceptionMessage(exception),
              location,
              List.of(formula, exception.getClass().getSimpleName())));
    }
  }
}
