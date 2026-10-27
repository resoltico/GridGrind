package dev.erst.gridgrind.engine.formulas;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.model.formula.FormulaCapabilityKind;
import dev.erst.gridgrind.core.model.formula.FormulaCellTarget;
import dev.erst.gridgrind.core.model.formula.FormulaMissingWorkbookPolicy;
import dev.erst.gridgrind.core.model.formula.FormulaPatternReport;
import dev.erst.gridgrind.core.model.formula.SheetFormulaSurfaceReport;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.engine.backend.xssf.CellNotFoundException;
import dev.erst.gridgrind.engine.backend.xssf.InvalidCellAddressException;
import dev.erst.gridgrind.engine.backend.xssf.SheetCollectionAdapter;
import dev.erst.gridgrind.engine.cells.CellErrors;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.jspecify.annotations.Nullable;

/** Owns workbook formula evaluation, cache writes, capabilities and stored formula facts. */
@SuppressWarnings({"PMD.ExcessiveImports", "PMD.CommentRequired"})
public final class FormulaAdapter {
  private final org.apache.poi.xssf.usermodel.XSSFWorkbook workbook;
  private final FormulaRuntime formulaRuntime;
  private final Runnable onMutation;

  public FormulaAdapter(
      org.apache.poi.xssf.usermodel.XSSFWorkbook workbook,
      FormulaRuntime formulaRuntime,
      Runnable onMutation) {
    this.workbook = Objects.requireNonNull(workbook, "workbook must not be null");
    this.formulaRuntime = Objects.requireNonNull(formulaRuntime, "formulaRuntime must not be null");
    this.onMutation = Objects.requireNonNull(onMutation, "onMutation must not be null");
  }

  public boolean recalculateOnOpenEnabled() {
    return workbook.getForceFormulaRecalculation();
  }

  public FormulaAdapter evaluateAll() {
    formulaRuntime.clearCachedResults();
    for (Sheet sheet : workbook) {
      for (Row row : sheet) {
        for (Cell cell : row) {
          if (cell.getCellType() == CellType.FORMULA) {
            evaluateFormulaCell(sheet.getSheetName(), cell);
          }
        }
      }
    }
    onMutation.run();
    return this;
  }

  public List<CalculationReport.FormulaCapability> assessAllCapabilities() {
    formulaRuntime.clearCachedResults();
    List<CalculationReport.FormulaCapability> assessments = new ArrayList<>();
    for (Sheet sheet : workbook) {
      for (Row row : sheet) {
        for (Cell cell : row) {
          if (cell.getCellType() == CellType.FORMULA) {
            assessments.add(assessFormulaCell(sheet.getSheetName(), cell));
          }
        }
      }
    }
    formulaRuntime.clearCachedResults();
    return List.copyOf(assessments);
  }

  public FormulaAdapter evaluate(List<FormulaCellTarget> cells) {
    Objects.requireNonNull(cells, "cells must not be null");
    formulaRuntime.clearCachedResults();
    for (FormulaCellTarget target : cells) {
      Objects.requireNonNull(target, "cells must not contain nulls");
      Cell cell = requiredCell(target.sheetName().value(), target.address().value());
      if (cell.getCellType() != CellType.FORMULA) {
        throw new IllegalArgumentException(
            "Cell "
                + target.sheetName().value()
                + "!"
                + target.address().value()
                + " is not a formula cell");
      }
      evaluateFormulaCell(target.sheetName().value(), cell);
    }
    onMutation.run();
    return this;
  }

  public List<CalculationReport.FormulaCapability> assessCapabilities(
      List<FormulaCellTarget> cells) {
    Objects.requireNonNull(cells, "cells must not be null");
    formulaRuntime.clearCachedResults();
    List<CalculationReport.FormulaCapability> assessments = new ArrayList<>(cells.size());
    for (FormulaCellTarget target : cells) {
      Objects.requireNonNull(target, "cells must not contain nulls");
      Cell cell = requiredCell(target.sheetName().value(), target.address().value());
      if (cell.getCellType() != CellType.FORMULA) {
        throw new IllegalArgumentException(
            "Cell "
                + target.sheetName().value()
                + "!"
                + target.address().value()
                + " is not a formula cell");
      }
      assessments.add(assessFormulaCell(target.sheetName().value(), cell));
    }
    formulaRuntime.clearCachedResults();
    return List.copyOf(assessments);
  }

  public FormulaAdapter clearCaches() {
    clearPersistedFormulaCaches();
    formulaRuntime.clearCachedResults();
    onMutation.run();
    return this;
  }

  public FormulaAdapter markRecalculateOnOpen() {
    workbook.setForceFormulaRecalculation(true);
    onMutation.run();
    return this;
  }

  private void evaluateFormulaCell(String sheetName, Cell cell) {
    try {
      formulaRuntime.evaluateFormulaCell(cell);
    } catch (RuntimeException exception) {
      throw FormulaExceptions.wrap(
          formulaRuntime,
          sheetName,
          new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(),
          cell.getCellFormula(),
          exception);
    }
  }

  private CalculationReport.FormulaCapability assessFormulaCell(String sheetName, Cell cell) {
    String address = new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString();
    String formula = cell.getCellFormula();
    try {
      CellValue evaluated = formulaRuntime.evaluate(cell);
      if (isCircularReference(evaluated)) {
        return new CalculationReport.FormulaCapability(
            new FormulaCellTarget(new SheetName(sheetName), new CellAddress(address)),
            formula,
            FormulaCapabilityKind.UNEVALUABLE_NOW,
            java.util.Optional.of(GridGrindProblemCode.CIRCULAR_FORMULA_REFERENCE),
            java.util.Optional.of("Formula evaluation detected a circular reference."));
      }
      return new CalculationReport.FormulaCapability(
          new FormulaCellTarget(new SheetName(sheetName), new CellAddress(address)),
          formula,
          FormulaCapabilityKind.EVALUABLE_NOW,
          java.util.Optional.empty(),
          java.util.Optional.empty());
    } catch (RuntimeException exception) {
      RuntimeException wrapped =
          FormulaExceptions.wrap(formulaRuntime, sheetName, address, formula, exception);
      return switch (wrapped) {
        case InvalidFormulaException invalid ->
            new CalculationReport.FormulaCapability(
                new FormulaCellTarget(
                    new SheetName(invalid.sheetName()), new CellAddress(invalid.address())),
                invalid.formula(),
                FormulaCapabilityKind.UNPARSEABLE_BY_POI,
                java.util.Optional.of(GridGrindProblemCode.INVALID_FORMULA),
                java.util.Optional.ofNullable(invalid.getMessage()));
        case MissingExternalWorkbookException missing ->
            new CalculationReport.FormulaCapability(
                new FormulaCellTarget(
                    new SheetName(missing.sheetName()), new CellAddress(missing.address())),
                missing.formula(),
                FormulaCapabilityKind.UNEVALUABLE_NOW,
                java.util.Optional.of(GridGrindProblemCode.MISSING_EXTERNAL_WORKBOOK),
                java.util.Optional.ofNullable(missing.getMessage()));
        case UnregisteredUserDefinedFunctionException unregistered ->
            new CalculationReport.FormulaCapability(
                new FormulaCellTarget(
                    new SheetName(unregistered.sheetName()),
                    new CellAddress(unregistered.address())),
                unregistered.formula(),
                FormulaCapabilityKind.UNEVALUABLE_NOW,
                java.util.Optional.of(GridGrindProblemCode.UNREGISTERED_USER_DEFINED_FUNCTION),
                java.util.Optional.ofNullable(unregistered.getMessage()));
        case UnsupportedFormulaException unsupported ->
            new CalculationReport.FormulaCapability(
                new FormulaCellTarget(
                    new SheetName(unsupported.sheetName()), new CellAddress(unsupported.address())),
                unsupported.formula(),
                FormulaCapabilityKind.UNEVALUABLE_NOW,
                java.util.Optional.of(GridGrindProblemCode.UNSUPPORTED_FORMULA),
                java.util.Optional.ofNullable(unsupported.getMessage()));
        default -> throw wrapped;
      };
    }
  }

  private static boolean isCircularReference(CellValue evaluated) {
    return evaluated != null
        && evaluated.getCellType() == CellType.ERROR
        && FormulaError.forInt(evaluated.getErrorValue()) == FormulaError.CIRCULAR_REF;
  }

  private void clearPersistedFormulaCaches() {
    for (Sheet sheet : workbook) {
      for (Row row : sheet) {
        for (Cell cell : row) {
          if (cell.getCellType() == CellType.FORMULA) {
            clearPersistedFormulaCache(cell);
          }
        }
      }
    }
  }

  private static void clearPersistedFormulaCache(Cell cell) {
    XSSFCell xssfCell = (XSSFCell) cell;
    var ctCell = xssfCell.getCTCell();
    if (ctCell.isSetV()) {
      ctCell.unsetV();
    }
    if (ctCell.isSetIs()) {
      ctCell.unsetIs();
    }
    if (ctCell.isSetT()) {
      ctCell.unsetT();
    }
  }

  // Request-local insertion order preserves physical formula/address order; report values are
  // allocated per cell/pattern.
  @SuppressWarnings({"PMD.AvoidInstantiatingObjectsInLoops", "PMD.UseConcurrentHashMap"})
  public static SheetFormulaSurfaceReport surface(Sheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Map<String, List<CellAddress>> formulas = new LinkedHashMap<>();
    int count = 0;
    for (Row row : sheet) {
      for (Cell cell : row) {
        if (cell.getCellType() == CellType.FORMULA) {
          formulas
              .computeIfAbsent(cell.getCellFormula(), _ -> new ArrayList<>())
              .add(
                  new CellAddress(
                      new CellReference(cell.getRowIndex(), cell.getColumnIndex())
                          .formatAsString()));
          count++;
        }
      }
    }
    List<FormulaPatternReport> patterns = new ArrayList<>(formulas.size());
    for (Map.Entry<String, List<CellAddress>> entry : formulas.entrySet()) {
      patterns.add(
          new FormulaPatternReport(entry.getKey(), entry.getValue().size(), entry.getValue()));
    }
    return new SheetFormulaSurfaceReport(
        new SheetLabel(sheet.getSheetName()), count, patterns.size(), patterns);
  }

  private Cell requiredCell(String sheetName, String address) {
    CellReference reference = parseCellReference(address);
    Sheet sheet = SheetCollectionAdapter.requiredSheet(workbook, sheetName);
    Row row = sheet.getRow(reference.getRow());
    if (row == null) {
      throw new CellNotFoundException(address);
    }
    Cell cell = row.getCell(reference.getCol());
    if (cell == null) {
      throw new CellNotFoundException(address);
    }
    return cell;
  }

  private static CellReference parseCellReference(String address) {
    try {
      return new CellReference(address);
    } catch (IllegalArgumentException exception) {
      throw new InvalidCellAddressException(address, exception);
    }
  }

  public static List<AnalysisFindingReport> formulaHealthFindings(
      String sheetName, org.apache.poi.ss.usermodel.Sheet sheet, FormulaRuntime formulaRuntime) {
    List<AnalysisFindingReport> findings = new ArrayList<>();
    FormulaContext formulaContext = formulaRuntime.context();
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
      List<AnalysisFindingReport> findings,
      String sheetName,
      Cell cell,
      FormulaRuntime formulaRuntime,
      FormulaContext formulaContext) {
    String address = cell.getAddress().formatAsString();
    String formula = cell.getCellFormula();
    AnalysisLocationReport.Cell location =
        new AnalysisLocationReport.Cell(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName), address);
    boolean missingExternalWorkbook =
        addExternalWorkbookFindings(findings, formulaContext, formula, location);
    addUserDefinedFunctionFinding(findings, formulaContext, formula, location);
    addVolatileFunctionFindings(findings, formula, location);
    addEvaluationFindings(
        findings, formulaRuntime, formulaContext, cell, formula, location, missingExternalWorkbook);
  }

  private static boolean addExternalWorkbookFindings(
      List<AnalysisFindingReport> findings,
      FormulaContext formulaContext,
      String formula,
      AnalysisLocationReport.Cell location) {
    boolean missingExternalWorkbook = false;
    for (String workbookName : FormulaExceptions.externalWorkbookNames(formula)) {
      if (formulaContext.hasExternalWorkbookBinding(workbookName)) {
        continue;
      }
      if (formulaContext.missingWorkbookPolicy() == FormulaMissingWorkbookPolicy.USE_CACHED_VALUE) {
        findings.add(cachedExternalWorkbookFinding(formula, workbookName, location));
      } else {
        missingExternalWorkbook = true;
        findings.add(missingExternalWorkbookFinding(formula, workbookName, location));
      }
    }
    return missingExternalWorkbook;
  }

  private static AnalysisFindingReport cachedExternalWorkbookFinding(
      String formula, String workbookName, AnalysisLocationReport.Cell location) {
    return new AnalysisFindingReport(
        AnalysisFindingCode.FORMULA_USES_CACHED_EXTERNAL_VALUE,
        AnalysisSeverity.INFO,
        "External workbook uses cached result",
        "Formula references external workbook "
            + workbookName
            + " and will fall back to the cached formula result when the workbook is missing.",
        location,
        List.of(formula, workbookName));
  }

  private static AnalysisFindingReport missingExternalWorkbookFinding(
      String formula, String workbookName, AnalysisLocationReport.Cell location) {
    return new AnalysisFindingReport(
        AnalysisFindingCode.FORMULA_MISSING_EXTERNAL_WORKBOOK,
        AnalysisSeverity.ERROR,
        "External workbook is missing or unbound",
        "Formula references external workbook " + workbookName + " but no binding is configured.",
        location,
        List.of(formula, workbookName));
  }

  private static void addUserDefinedFunctionFinding(
      List<AnalysisFindingReport> findings,
      FormulaContext formulaContext,
      String formula,
      AnalysisLocationReport.Cell location) {
    String leadingFunctionName = FormulaExceptions.leadingFunctionName(formula).orElse(null);
    if (!FormulaAdapter.hasUnregisteredUserDefinedFunction(
        formulaContext, leadingFunctionName, formula)) {
      return;
    }
    findings.add(
        new AnalysisFindingReport(
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
      List<AnalysisFindingReport> findings, String formula, AnalysisLocationReport.Cell location) {
    FormulaAdapter.volatileFunctions(formula)
        .forEach(
            functionName ->
                findings.add(
                    new AnalysisFindingReport(
                        AnalysisFindingCode.FORMULA_VOLATILE_FUNCTION,
                        AnalysisSeverity.INFO,
                        "Volatile formula function",
                        "Formula uses volatile function " + functionName + ".",
                        location,
                        List.of(formula, functionName))));
  }

  private static void addEvaluationFindings(
      List<AnalysisFindingReport> findings,
      FormulaRuntime formulaRuntime,
      FormulaContext formulaContext,
      Cell cell,
      String formula,
      AnalysisLocationReport.Cell location,
      boolean missingExternalWorkbook) {
    try {
      var evaluated = formulaRuntime.evaluate(cell);
      if (evaluated != null && evaluated.getCellType() == CellType.ERROR) {
        findings.add(
            new AnalysisFindingReport(
                AnalysisFindingCode.FORMULA_ERROR_RESULT,
                AnalysisSeverity.ERROR,
                "Formula evaluates to an error",
                "Formula currently evaluates to "
                    + CellErrors.toReportedWireValue(FormulaError.forInt(evaluated.getErrorValue()))
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
      List<AnalysisFindingReport> findings,
      FormulaContext formulaContext,
      String formula,
      AnalysisLocationReport.Cell location,
      boolean missingExternalWorkbook,
      RuntimeException exception) {
    if (FormulaExceptions.isMissingExternalWorkbookFailure(exception)) {
      if (!missingExternalWorkbook) {
        String workbookName = FormulaExceptions.missingExternalWorkbookName(exception, formula);
        findings.add(
            new AnalysisFindingReport(
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
          new AnalysisFindingReport(
              AnalysisFindingCode.FORMULA_EVALUATION_FAILURE,
              AnalysisSeverity.ERROR,
              "Formula evaluation failed",
              "Formula evaluation failed: " + FormulaAdapter.exceptionMessage(exception),
              location,
              List.of(formula, exception.getClass().getSimpleName())));
    }
  }

  public static int formulaCellCount(org.apache.poi.ss.usermodel.Sheet sheet) {
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

  public static boolean containsExternalWorkbookReference(String formula) {
    return formula.contains("[") && formula.contains("]");
  }

  public static List<String> volatileFunctions(String formula) {
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

  public static boolean hasUnregisteredUserDefinedFunction(
      FormulaContext formulaContext, @Nullable String leadingFunctionName, String formula) {
    return leadingFunctionName != null
        && FormulaExceptions.isUnregisteredUserDefinedFunctionFailure(
            formulaContext,
            new org.apache.poi.ss.formula.eval.NotImplementedFunctionException(leadingFunctionName),
            formula);
  }

  public static String exceptionMessage(Exception exception) {
    if (exception.getMessage() == null) {
      return exception.getClass().getSimpleName();
    }
    return exception.getMessage();
  }

  /** Reads stored formula presence without evaluating or changing cached results. */
  public boolean containsFormula() {
    for (Sheet sheet : workbook) {
      for (Row row : sheet) {
        for (Cell cell : row) {
          if (cell.getCellType() == CellType.FORMULA) return true;
        }
      }
    }
    return false;
  }

  /** Checks whether the exact stored formula still owns an authored origin. */
  public boolean matchesStoredFormula(String sheetName, String address, String formula) {
    Sheet sheet = workbook.getSheet(sheetName);
    if (sheet == null) return false;
    CellReference reference = new CellReference(address);
    Row row = sheet.getRow(reference.getRow());
    Cell cell = row == null ? null : row.getCell(reference.getCol());
    return cell != null
        && cell.getCellType() == CellType.FORMULA
        && formula.equals(cell.getCellFormula());
  }
}
