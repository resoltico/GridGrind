package dev.erst.gridgrind.excel;

import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;

/** Derived-analysis support for formula and hyperlink health on one sheet. */
final class ExcelSheetAnalysisSupport {
  private final org.apache.poi.ss.usermodel.Sheet sheet;
  private final ExcelFormulaRuntime formulaRuntime;

  ExcelSheetAnalysisSupport(
      org.apache.poi.ss.usermodel.Sheet sheet, ExcelFormulaRuntime formulaRuntime) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.formulaRuntime = Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");
  }

  int formulaCellCount() {
    int count = 0;
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (cell.getCellType() == CellType.FORMULA) {
          count++;
        }
      }
    }
    return count;
  }

  List<WorkbookAnalysis.AnalysisFinding> formulaHealthFindings(String sheetName) {
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    ExcelFormulaRuntimeContext formulaContext = formulaRuntime.context();
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (cell.getCellType() != CellType.FORMULA) {
          continue;
        }
        String address = cell.getAddress().formatAsString();
        String formula = cell.getCellFormula();
        WorkbookAnalysis.AnalysisLocation.Cell location = cellLocation(sheetName, address);
        List<String> externalWorkbookNames = FormulaExceptions.externalWorkbookNames(formula);
        boolean hasStaticMissingExternalWorkbookFinding = false;
        for (String workbookName : externalWorkbookNames) {
          if (formulaContext.hasExternalWorkbookBinding(workbookName)) {
            continue;
          }
          if (formulaContext.missingWorkbookPolicy()
              == ExcelFormulaMissingWorkbookPolicy.USE_CACHED_VALUE) {
            findings.add(
                new WorkbookAnalysis.AnalysisFinding(
                    WorkbookAnalysis.AnalysisFindingCode.FORMULA_USES_CACHED_EXTERNAL_VALUE,
                    WorkbookAnalysis.AnalysisSeverity.INFO,
                    "External workbook uses cached result",
                    "Formula references external workbook "
                        + workbookName
                        + " and will fall back to the cached formula result when the workbook is missing.",
                    location,
                    List.of(formula, workbookName)));
          } else {
            hasStaticMissingExternalWorkbookFinding = true;
            findings.add(
                new WorkbookAnalysis.AnalysisFinding(
                    WorkbookAnalysis.AnalysisFindingCode.FORMULA_MISSING_EXTERNAL_WORKBOOK,
                    WorkbookAnalysis.AnalysisSeverity.ERROR,
                    "External workbook is missing or unbound",
                    "Formula references external workbook "
                        + workbookName
                        + " but no binding is configured.",
                    location,
                    List.of(formula, workbookName)));
          }
        }
        String leadingFunctionName = FormulaExceptions.leadingFunctionName(formula).orElse(null);
        if (hasUnregisteredUserDefinedFunction(formulaContext, leadingFunctionName, formula)) {
          findings.add(
              new WorkbookAnalysis.AnalysisFinding(
                  WorkbookAnalysis.AnalysisFindingCode.FORMULA_UNREGISTERED_USER_DEFINED_FUNCTION,
                  WorkbookAnalysis.AnalysisSeverity.ERROR,
                  "User-defined function is not registered",
                  "Formula references user-defined function "
                      + leadingFunctionName
                      + " but the current formula environment does not register it.",
                  location,
                  List.of(formula, leadingFunctionName)));
        }
        volatileFunctions(formula)
            .forEach(
                functionName ->
                    findings.add(
                        new WorkbookAnalysis.AnalysisFinding(
                            WorkbookAnalysis.AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                            WorkbookAnalysis.AnalysisSeverity.INFO,
                            "Volatile formula function",
                            "Formula uses volatile function " + functionName + ".",
                            location,
                            List.of(formula, functionName))));
        try {
          org.apache.poi.ss.usermodel.CellValue evaluated = formulaRuntime.evaluate(cell);
          if (evaluated != null && evaluated.getCellType() == CellType.ERROR) {
            findings.add(
                new WorkbookAnalysis.AnalysisFinding(
                    WorkbookAnalysis.AnalysisFindingCode.FORMULA_ERROR_RESULT,
                    WorkbookAnalysis.AnalysisSeverity.ERROR,
                    "Formula evaluates to an error",
                    "Formula currently evaluates to "
                        + FormulaError.forInt(evaluated.getErrorValue()).getString()
                        + ".",
                    location,
                    List.of(formula)));
          }
        } catch (RuntimeException exception) {
          if (FormulaExceptions.isMissingExternalWorkbookFailure(exception)) {
            if (!hasStaticMissingExternalWorkbookFinding) {
              String workbookName =
                  FormulaExceptions.missingExternalWorkbookName(exception, formula);
              findings.add(
                  new WorkbookAnalysis.AnalysisFinding(
                      WorkbookAnalysis.AnalysisFindingCode.FORMULA_MISSING_EXTERNAL_WORKBOOK,
                      WorkbookAnalysis.AnalysisSeverity.ERROR,
                      "External workbook is missing or unbound",
                      "Formula evaluation failed because external workbook "
                          + workbookName
                          + " could not be resolved.",
                      location,
                      List.of(formula, workbookName == null ? "" : workbookName)));
            }
          } else if (!FormulaExceptions.isUnregisteredUserDefinedFunctionFailure(
              formulaContext, exception, formula)) {
            findings.add(
                new WorkbookAnalysis.AnalysisFinding(
                    WorkbookAnalysis.AnalysisFindingCode.FORMULA_EVALUATION_FAILURE,
                    WorkbookAnalysis.AnalysisSeverity.ERROR,
                    "Formula evaluation failed",
                    "Formula evaluation failed: " + exceptionMessage(exception),
                    location,
                    List.of(formula, exception.getClass().getSimpleName())));
          }
        }
      }
    }
    return List.copyOf(findings);
  }

  int hyperlinkCount() {
    int count = 0;
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (hasUsableHyperlink(cell.getHyperlink())) {
          count++;
        }
      }
    }
    return count;
  }

  List<WorkbookAnalysis.AnalysisFinding> hyperlinkHealthFindings(
      WorkbookLocation workbookLocation) {
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (Row row : sheet) {
      for (Cell cell : row) {
        org.apache.poi.ss.usermodel.Hyperlink hyperlink = cell.getHyperlink();
        if (!hasUsableHyperlink(hyperlink)) {
          continue;
        }
        HyperlinkType hyperlinkType = hyperlink.getType();
        String address = cell.getAddress().formatAsString();
        WorkbookAnalysis.AnalysisLocation.Cell location =
            cellLocation(sheet.getSheetName(), address);
        findings.addAll(
            hyperlinkTargetFindings(
                location, hyperlinkType, hyperlink.getAddress(), workbookLocation));
      }
    }
    return List.copyOf(findings);
  }

  List<WorkbookAnalysis.AnalysisFinding> hyperlinkHealthFindings() {
    return hyperlinkHealthFindings(new WorkbookLocation.UnsavedWorkbook());
  }

  static List<WorkbookAnalysis.AnalysisFinding> externalHyperlinkFindings(
      WorkbookAnalysis.AnalysisLocation.Cell location,
      String target,
      String expectedShape,
      boolean valid) {
    Objects.requireNonNull(location, "location must not be null");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(expectedShape, "expectedShape must not be null");
    return valid
        ? List.of()
        : List.of(
            malformedHyperlinkFinding(
                location,
                "target does not match the expected " + expectedShape + " shape",
                List.of(target)));
  }

  List<WorkbookAnalysis.AnalysisFinding> hyperlinkTargetFindings(
      WorkbookAnalysis.AnalysisLocation.Cell location,
      HyperlinkType hyperlinkType,
      String target,
      WorkbookLocation workbookLocation) {
    Objects.requireNonNull(location, "location must not be null");
    Objects.requireNonNull(hyperlinkType, "hyperlinkType must not be null");
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");

    if (hasMissingHyperlinkTarget(target)) {
      return List.of(
          malformedHyperlinkFinding(
              location, "Hyperlink target is blank or missing.", List.of(hyperlinkType.name())));
    }
    return switch (hyperlinkType) {
      case URL ->
          externalHyperlinkFindings(
              location,
              target,
              ExcelHyperlinkType.URL.name(),
              ExcelHyperlinkValidation.isValidUrlTarget(target));
      case EMAIL ->
          externalHyperlinkFindings(
              location,
              target,
              ExcelHyperlinkType.EMAIL.name(),
              ExcelHyperlinkValidation.isValidEmailTarget(target));
      case FILE -> fileHyperlinkFindings(location, target, workbookLocation);
      case DOCUMENT -> {
        List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
        validateDocumentHyperlinkTarget(location, target, findings);
        yield List.copyOf(findings);
      }
      case NONE -> List.of();
    };
  }

  void validateDocumentHyperlinkTarget(
      WorkbookAnalysis.AnalysisLocation.Cell location,
      String target,
      List<WorkbookAnalysis.AnalysisFinding> findings) {
    int bangIndex = target.indexOf('!');
    if (bangIndex <= 0 || bangIndex >= target.length() - 1) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET,
              WorkbookAnalysis.AnalysisSeverity.ERROR,
              "Invalid document hyperlink target",
              "Document hyperlink target must include a sheet and cell or range reference.",
              location,
              List.of(target)));
      return;
    }

    String targetSheetName = unquoteSheetName(target.substring(0, bangIndex));
    String targetAddress = target.substring(bangIndex + 1);
    if (sheet.getWorkbook().getSheet(targetSheetName) == null) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_MISSING_DOCUMENT_SHEET,
              WorkbookAnalysis.AnalysisSeverity.ERROR,
              "Document hyperlink targets a missing sheet",
              "Document hyperlink target sheet does not exist: " + targetSheetName,
              location,
              List.of(target)));
      return;
    }

    try {
      if (targetAddress.contains(":")) {
        ExcelRange.parse(targetAddress);
      } else {
        new CellReference(targetAddress);
      }
    } catch (IllegalArgumentException exception) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_INVALID_DOCUMENT_TARGET,
              WorkbookAnalysis.AnalysisSeverity.ERROR,
              "Invalid document hyperlink target",
              "Document hyperlink target is not a valid cell or range reference.",
              location,
              List.of(target)));
    }
  }

  static List<WorkbookAnalysis.AnalysisFinding> fileHyperlinkFindings(
      WorkbookAnalysis.AnalysisLocation.Cell location,
      String target,
      WorkbookLocation workbookLocation) {
    Objects.requireNonNull(location, "location must not be null");
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(workbookLocation, "workbookLocation must not be null");

    return switch (ExcelFileHyperlinkTargets.resolve(target, workbookLocation)) {
      case FileHyperlinkResolution.MalformedPath malformedPath ->
          List.of(
              malformedHyperlinkFinding(
                  location, malformedPath.reason(), List.of(malformedPath.path())));
      case FileHyperlinkResolution.UnresolvedRelativePath unresolvedRelativePath ->
          List.of(
              new WorkbookAnalysis.AnalysisFinding(
                  WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_UNRESOLVED_FILE_TARGET,
                  WorkbookAnalysis.AnalysisSeverity.WARNING,
                  "Relative file hyperlink cannot be resolved yet",
                  "Relative file hyperlink targets cannot be validated until the workbook has a"
                      + " filesystem location.",
                  location,
                  List.of(unresolvedRelativePath.path())));
      case FileHyperlinkResolution.ResolvedPath resolvedPath ->
          Files.notExists(resolvedPath.resolvedPath())
              ? List.of(
                  new WorkbookAnalysis.AnalysisFinding(
                      WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_MISSING_FILE_TARGET,
                      WorkbookAnalysis.AnalysisSeverity.ERROR,
                      "File hyperlink targets a missing path",
                      missingFileTargetMessage(resolvedPath, workbookLocation),
                      location,
                      missingFileTargetEvidence(resolvedPath, workbookLocation)))
              : List.of();
    };
  }

  private static String missingFileTargetMessage(
      FileHyperlinkResolution.ResolvedPath resolvedPath, WorkbookLocation workbookLocation) {
    if (workbookLocation instanceof WorkbookLocation.StoredWorkbook storedWorkbook
        && ExcelFileHyperlinkTargets.isRelativeStoredPath(resolvedPath.path())) {
      return "Resolved relative file hyperlink path does not exist against workbook directory "
          + storedWorkbook.baseDirectory().orElseThrow()
          + ": "
          + resolvedPath.resolvedPath();
    }
    return "Resolved file hyperlink path does not exist: " + resolvedPath.resolvedPath();
  }

  private static List<String> missingFileTargetEvidence(
      FileHyperlinkResolution.ResolvedPath resolvedPath, WorkbookLocation workbookLocation) {
    if (workbookLocation instanceof WorkbookLocation.StoredWorkbook storedWorkbook
        && ExcelFileHyperlinkTargets.isRelativeStoredPath(resolvedPath.path())) {
      return List.of(
          resolvedPath.path(),
          storedWorkbook.baseDirectory().orElseThrow().toString(),
          resolvedPath.resolvedPath().toString());
    }
    return List.of(resolvedPath.path(), resolvedPath.resolvedPath().toString());
  }

  static boolean hasUsableHyperlink(org.apache.poi.ss.usermodel.Hyperlink hyperlink) {
    return hyperlink != null && hasUsableHyperlinkType(hyperlink.getType());
  }

  static boolean hasUsableHyperlinkType(HyperlinkType hyperlinkType) {
    return hyperlinkType != null && hyperlinkType != HyperlinkType.NONE;
  }

  static boolean hasMissingHyperlinkTarget(String target) {
    return target == null || target.isBlank();
  }

  static String unquoteSheetName(String sheetName) {
    if (!sheetName.startsWith("'")) {
      return sheetName;
    }
    if (!sheetName.endsWith("'")) {
      return sheetName;
    }
    if (sheetName.length() < 2) {
      return sheetName;
    }
    return sheetName.substring(1, sheetName.length() - 1).replace("''", "'");
  }

  static boolean containsExternalWorkbookReference(String formula) {
    return formula.contains("[") && formula.contains("]");
  }

  static List<String> volatileFunctions(String formula) {
    String upper = formula.toUpperCase(Locale.ROOT);
    List<String> functions = new ArrayList<>();
    for (String candidate :
        List.of("NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT", "CELL", "INFO")) {
      if (upper.contains(candidate + "(")) {
        functions.add(candidate);
      }
    }
    return List.copyOf(functions);
  }

  private static WorkbookAnalysis.AnalysisFinding malformedHyperlinkFinding(
      WorkbookAnalysis.AnalysisLocation.Cell location, String message, List<String> evidence) {
    return new WorkbookAnalysis.AnalysisFinding(
        WorkbookAnalysis.AnalysisFindingCode.HYPERLINK_MALFORMED_TARGET,
        WorkbookAnalysis.AnalysisSeverity.ERROR,
        "Malformed hyperlink target",
        "Hyperlink target is malformed: " + message,
        location,
        evidence);
  }

  private static WorkbookAnalysis.AnalysisLocation.Cell cellLocation(
      String sheetName, String address) {
    return new WorkbookAnalysis.AnalysisLocation.Cell(sheetName, address);
  }

  private static boolean hasUnregisteredUserDefinedFunction(
      ExcelFormulaRuntimeContext formulaContext, String leadingFunctionName, String formula) {
    return leadingFunctionName != null
        && FormulaExceptions.isUnregisteredUserDefinedFunctionFailure(
            formulaContext,
            new org.apache.poi.ss.formula.eval.NotImplementedFunctionException(leadingFunctionName),
            formula);
  }

  static String exceptionMessage(Exception exception) {
    if (exception.getMessage() == null) {
      return exception.getClass().getSimpleName();
    }
    return exception.getMessage();
  }
}
