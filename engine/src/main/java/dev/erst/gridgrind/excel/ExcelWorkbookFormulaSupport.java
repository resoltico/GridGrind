package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.formula.CalculationReport;
import dev.erst.gridgrind.core.model.formula.FormulaCapabilityKind;
import dev.erst.gridgrind.core.model.formula.FormulaCellTarget;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;

/**
 * Formula evaluation, capability assessment, and cache mutation support for {@link ExcelWorkbook}.
 */
final class ExcelWorkbookFormulaSupport {
  private ExcelWorkbookFormulaSupport() {}

  static ExcelWorkbook evaluateAllFormulas(ExcelWorkbook workbook) {
    workbook.context().formulaRuntime().clearCachedResults();
    for (Sheet sheet : workbook.context().workbook()) {
      for (Row row : sheet) {
        for (Cell cell : row) {
          if (cell.getCellType() == CellType.FORMULA) {
            evaluateFormulaCell(workbook, sheet.getSheetName(), cell);
          }
        }
      }
    }
    workbook.markPackageMutated();
    return workbook;
  }

  static List<CalculationReport.FormulaCapability> assessAllFormulaCapabilities(
      ExcelWorkbook workbook) {
    workbook.context().formulaRuntime().clearCachedResults();
    List<CalculationReport.FormulaCapability> assessments = new ArrayList<>();
    for (Sheet sheet : workbook.context().workbook()) {
      for (Row row : sheet) {
        for (Cell cell : row) {
          if (cell.getCellType() == CellType.FORMULA) {
            assessments.add(assessFormulaCell(workbook, sheet.getSheetName(), cell));
          }
        }
      }
    }
    workbook.context().formulaRuntime().clearCachedResults();
    return List.copyOf(assessments);
  }

  static ExcelWorkbook evaluateFormulaCells(ExcelWorkbook workbook, List<FormulaCellTarget> cells) {
    Objects.requireNonNull(cells, "cells must not be null");
    workbook.context().formulaRuntime().clearCachedResults();
    for (FormulaCellTarget target : cells) {
      Objects.requireNonNull(target, "cells must not contain nulls");
      Cell cell = workbook.requiredCell(target.sheetName().value(), target.address().value());
      if (cell.getCellType() != CellType.FORMULA) {
        throw new IllegalArgumentException(
            "Cell "
                + target.sheetName().value()
                + "!"
                + target.address().value()
                + " is not a formula cell");
      }
      evaluateFormulaCell(workbook, target.sheetName().value(), cell);
    }
    workbook.markPackageMutated();
    return workbook;
  }

  static List<CalculationReport.FormulaCapability> assessFormulaCellCapabilities(
      ExcelWorkbook workbook, List<FormulaCellTarget> cells) {
    Objects.requireNonNull(cells, "cells must not be null");
    workbook.context().formulaRuntime().clearCachedResults();
    List<CalculationReport.FormulaCapability> assessments = new ArrayList<>(cells.size());
    for (FormulaCellTarget target : cells) {
      Objects.requireNonNull(target, "cells must not contain nulls");
      Cell cell = workbook.requiredCell(target.sheetName().value(), target.address().value());
      if (cell.getCellType() != CellType.FORMULA) {
        throw new IllegalArgumentException(
            "Cell "
                + target.sheetName().value()
                + "!"
                + target.address().value()
                + " is not a formula cell");
      }
      assessments.add(assessFormulaCell(workbook, target.sheetName().value(), cell));
    }
    workbook.context().formulaRuntime().clearCachedResults();
    return List.copyOf(assessments);
  }

  static ExcelWorkbook clearFormulaCaches(ExcelWorkbook workbook) {
    clearPersistedFormulaCaches(workbook);
    workbook.context().formulaRuntime().clearCachedResults();
    workbook.markPackageMutated();
    return workbook;
  }

  static ExcelWorkbook forceFormulaRecalculationOnOpen(ExcelWorkbook workbook) {
    workbook.context().workbook().setForceFormulaRecalculation(true);
    workbook.markPackageMutated();
    return workbook;
  }

  private static void evaluateFormulaCell(ExcelWorkbook workbook, String sheetName, Cell cell) {
    try {
      workbook.context().formulaRuntime().evaluateFormulaCell(cell);
    } catch (RuntimeException exception) {
      throw FormulaExceptions.wrap(
          workbook.context().formulaRuntime(),
          sheetName,
          new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString(),
          cell.getCellFormula(),
          exception);
    }
  }

  private static CalculationReport.FormulaCapability assessFormulaCell(
      ExcelWorkbook workbook, String sheetName, Cell cell) {
    String address = new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString();
    String formula = cell.getCellFormula();
    try {
      CellValue evaluated = workbook.context().formulaRuntime().evaluate(cell);
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
          FormulaExceptions.wrap(
              workbook.context().formulaRuntime(), sheetName, address, formula, exception);
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

  private static void clearPersistedFormulaCaches(ExcelWorkbook workbook) {
    for (Sheet sheet : workbook.context().workbook()) {
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
}
