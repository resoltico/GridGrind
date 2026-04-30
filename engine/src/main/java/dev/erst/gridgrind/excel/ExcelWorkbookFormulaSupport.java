package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
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

  static List<ExcelFormulaCapabilityAssessment> assessAllFormulaCapabilities(
      ExcelWorkbook workbook) {
    workbook.context().formulaRuntime().clearCachedResults();
    List<ExcelFormulaCapabilityAssessment> assessments = new ArrayList<>();
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

  static ExcelWorkbook evaluateFormulaCells(
      ExcelWorkbook workbook, List<ExcelFormulaCellTarget> cells) {
    Objects.requireNonNull(cells, "cells must not be null");
    workbook.context().formulaRuntime().clearCachedResults();
    for (ExcelFormulaCellTarget target : cells) {
      Objects.requireNonNull(target, "cells must not contain nulls");
      Cell cell = workbook.requiredCell(target.sheetName(), target.address());
      if (cell.getCellType() != CellType.FORMULA) {
        throw new IllegalArgumentException(
            "Cell " + target.sheetName() + "!" + target.address() + " is not a formula cell");
      }
      evaluateFormulaCell(workbook, target.sheetName(), cell);
    }
    workbook.markPackageMutated();
    return workbook;
  }

  static List<ExcelFormulaCapabilityAssessment> assessFormulaCellCapabilities(
      ExcelWorkbook workbook, List<ExcelFormulaCellTarget> cells) {
    Objects.requireNonNull(cells, "cells must not be null");
    workbook.context().formulaRuntime().clearCachedResults();
    List<ExcelFormulaCapabilityAssessment> assessments = new ArrayList<>(cells.size());
    for (ExcelFormulaCellTarget target : cells) {
      Objects.requireNonNull(target, "cells must not contain nulls");
      Cell cell = workbook.requiredCell(target.sheetName(), target.address());
      if (cell.getCellType() != CellType.FORMULA) {
        throw new IllegalArgumentException(
            "Cell " + target.sheetName() + "!" + target.address() + " is not a formula cell");
      }
      assessments.add(assessFormulaCell(workbook, target.sheetName(), cell));
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

  private static ExcelFormulaCapabilityAssessment assessFormulaCell(
      ExcelWorkbook workbook, String sheetName, Cell cell) {
    String address = new CellReference(cell.getRowIndex(), cell.getColumnIndex()).formatAsString();
    String formula = cell.getCellFormula();
    try {
      workbook.context().formulaRuntime().evaluate(cell);
      return new ExcelFormulaCapabilityAssessment(
          sheetName, address, formula, ExcelFormulaCapabilityKind.EVALUABLE_NOW, null, null);
    } catch (RuntimeException exception) {
      RuntimeException wrapped =
          FormulaExceptions.wrap(
              workbook.context().formulaRuntime(), sheetName, address, formula, exception);
      return switch (wrapped) {
        case InvalidFormulaException invalid ->
            new ExcelFormulaCapabilityAssessment(
                invalid.sheetName(),
                invalid.address(),
                invalid.formula(),
                ExcelFormulaCapabilityKind.UNPARSEABLE_BY_POI,
                ExcelFormulaCapabilityIssue.INVALID_FORMULA,
                invalid.getMessage());
        case MissingExternalWorkbookException missing ->
            new ExcelFormulaCapabilityAssessment(
                missing.sheetName(),
                missing.address(),
                missing.formula(),
                ExcelFormulaCapabilityKind.UNEVALUABLE_NOW,
                ExcelFormulaCapabilityIssue.MISSING_EXTERNAL_WORKBOOK,
                missing.getMessage());
        case UnregisteredUserDefinedFunctionException unregistered ->
            new ExcelFormulaCapabilityAssessment(
                unregistered.sheetName(),
                unregistered.address(),
                unregistered.formula(),
                ExcelFormulaCapabilityKind.UNEVALUABLE_NOW,
                ExcelFormulaCapabilityIssue.UNREGISTERED_USER_DEFINED_FUNCTION,
                unregistered.getMessage());
        case UnsupportedFormulaException unsupported ->
            new ExcelFormulaCapabilityAssessment(
                unsupported.sheetName(),
                unsupported.address(),
                unsupported.formula(),
                ExcelFormulaCapabilityKind.UNEVALUABLE_NOW,
                ExcelFormulaCapabilityIssue.UNSUPPORTED_FORMULA,
                unsupported.getMessage());
        default -> throw wrapped;
      };
    }
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
