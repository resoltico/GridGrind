package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.FormulaParser;
import org.apache.poi.ss.formula.FormulaRenderer;
import org.apache.poi.ss.formula.FormulaShifter;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.ss.formula.ptg.Ptg;
import org.apache.poi.xssf.usermodel.XSSFEvaluationWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;

/** Normalizes sheet data validations across supported structural insert operations. */
final class ExcelDataValidationStructureSupport {
  private ExcelDataValidationStructureSupport() {}

  @SuppressWarnings("PMD.CloseResource")
  static List<CTDataValidation> expectedValidationsAfterInsertRows(
      XSSFSheet sheet, int rowIndex, int rowCount) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    if (!sheet.getCTWorksheet().isSetDataValidations()) {
      return List.of();
    }
    XSSFWorkbook workbook = sheet.getWorkbook();
    int sheetIndex = workbook.getSheetIndex(sheet);
    XSSFEvaluationWorkbook evaluationWorkbook = XSSFEvaluationWorkbook.create(workbook);
    FormulaShifter formulaShifter =
        FormulaShifter.createForRowShift(
            sheetIndex,
            sheet.getSheetName(),
            rowIndex,
            SpreadsheetVersion.EXCEL2007.getLastRowIndex(),
            rowCount,
            SpreadsheetVersion.EXCEL2007);
    List<CTDataValidation> rewritten = new ArrayList<>();
    for (CTDataValidation validation : copiedValidations(sheet)) {
      List<String> shiftedRanges =
          shiftedRangesForInsertedRows(sheet, validation, rowIndex, rowCount);
      if (shiftedRanges.isEmpty()) {
        continue;
      }
      validation.setSqref(shiftedRanges);
      shiftValidationFormulas(validation, evaluationWorkbook, formulaShifter, sheetIndex, sheet);
      rewritten.add(validation);
    }
    return List.copyOf(rewritten);
  }

  @SuppressWarnings("PMD.CloseResource")
  static List<CTDataValidation> expectedValidationsAfterInsertColumns(
      XSSFSheet sheet, int columnIndex, int columnCount) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    if (!sheet.getCTWorksheet().isSetDataValidations()) {
      return List.of();
    }
    XSSFWorkbook workbook = sheet.getWorkbook();
    int sheetIndex = workbook.getSheetIndex(sheet);
    XSSFEvaluationWorkbook evaluationWorkbook = XSSFEvaluationWorkbook.create(workbook);
    FormulaShifter formulaShifter =
        FormulaShifter.createForColumnShift(
            sheetIndex,
            sheet.getSheetName(),
            columnIndex,
            SpreadsheetVersion.EXCEL2007.getLastColumnIndex(),
            columnCount,
            SpreadsheetVersion.EXCEL2007);
    List<CTDataValidation> rewritten = new ArrayList<>();
    for (CTDataValidation validation : copiedValidations(sheet)) {
      List<String> shiftedRanges =
          shiftedRangesForInsertedColumns(sheet, validation, columnIndex, columnCount);
      if (shiftedRanges.isEmpty()) {
        continue;
      }
      validation.setSqref(shiftedRanges);
      shiftValidationFormulas(validation, evaluationWorkbook, formulaShifter, sheetIndex, sheet);
      rewritten.add(validation);
    }
    return List.copyOf(rewritten);
  }

  static void replaceDataValidations(XSSFSheet sheet, List<CTDataValidation> validations) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(validations, "validations must not be null");
    if (sheet.getCTWorksheet().isSetDataValidations()) {
      sheet.getCTWorksheet().unsetDataValidations();
    }
    if (validations.isEmpty()) {
      return;
    }
    CTDataValidations target = sheet.getCTWorksheet().addNewDataValidations();
    for (CTDataValidation validation : validations) {
      CTDataValidation copied = target.addNewDataValidation();
      copied.set(validation);
    }
    target.setCount(target.sizeOfDataValidationArray());
  }

  private static List<CTDataValidation> copiedValidations(XSSFSheet sheet) {
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    List<CTDataValidation> copied = new ArrayList<>(dataValidations.sizeOfDataValidationArray());
    for (CTDataValidation validation : dataValidations.getDataValidationArray()) {
      copied.add((CTDataValidation) validation.copy());
    }
    return List.copyOf(copied);
  }

  private static List<String> shiftedRangesForInsertedRows(
      XSSFSheet sheet, CTDataValidation validation, int rowIndex, int rowCount) {
    List<String> shifted = new ArrayList<>();
    for (String rawRange : ExcelSqrefSupport.normalizedSqref(validation.getSqref())) {
      ExcelRange range = ExcelRange.parse(rawRange);
      shifted.addAll(shiftedRangesForInsertedRows(sheet, rawRange, range, rowIndex, rowCount));
    }
    return List.copyOf(shifted);
  }

  private static List<String> shiftedRangesForInsertedColumns(
      XSSFSheet sheet, CTDataValidation validation, int columnIndex, int columnCount) {
    List<String> shifted = new ArrayList<>();
    for (String rawRange : ExcelSqrefSupport.normalizedSqref(validation.getSqref())) {
      ExcelRange range = ExcelRange.parse(rawRange);
      shifted.addAll(
          shiftedRangesForInsertedColumns(sheet, rawRange, range, columnIndex, columnCount));
    }
    return List.copyOf(shifted);
  }

  private static List<String> shiftedRangesForInsertedRows(
      XSSFSheet sheet, String rawRange, ExcelRange range, int rowIndex, int rowCount) {
    if (range.lastRow() < rowIndex) {
      return List.of(rawRange);
    }
    if (range.firstRow() >= rowIndex) {
      return List.of(
          shiftedRangeForRows(
              sheet,
              rawRange,
              range.firstRow() + rowCount,
              range.lastRow() + rowCount,
              range.firstColumn(),
              range.lastColumn()));
    }
    List<String> shifted = new ArrayList<>(2);
    shifted.add(
        ExcelSheetStructureSupport.formatRange(
            new ExcelRange(
                range.firstRow(), rowIndex - 1, range.firstColumn(), range.lastColumn())));
    shifted.add(
        shiftedRangeForRows(
            sheet,
            rawRange,
            rowIndex + rowCount,
            range.lastRow() + rowCount,
            range.firstColumn(),
            range.lastColumn()));
    return List.copyOf(shifted);
  }

  private static List<String> shiftedRangesForInsertedColumns(
      XSSFSheet sheet, String rawRange, ExcelRange range, int columnIndex, int columnCount) {
    if (range.lastColumn() < columnIndex) {
      return List.of(rawRange);
    }
    if (range.firstColumn() >= columnIndex) {
      return List.of(
          shiftedRangeForColumns(
              sheet,
              rawRange,
              range.firstRow(),
              range.lastRow(),
              range.firstColumn() + columnCount,
              range.lastColumn() + columnCount));
    }
    List<String> shifted = new ArrayList<>(2);
    shifted.add(
        ExcelSheetStructureSupport.formatRange(
            new ExcelRange(
                range.firstRow(), range.lastRow(), range.firstColumn(), columnIndex - 1)));
    shifted.add(
        shiftedRangeForColumns(
            sheet,
            rawRange,
            range.firstRow(),
            range.lastRow(),
            columnIndex + columnCount,
            range.lastColumn() + columnCount));
    return List.copyOf(shifted);
  }

  private static String shiftedRangeForRows(
      XSSFSheet sheet,
      String rawRange,
      int firstRow,
      int lastRow,
      int firstColumn,
      int lastColumn) {
    if (lastRow > ExcelRowSpan.MAX_ROW_INDEX) {
      throw new IllegalArgumentException(
          "INSERT_ROWS cannot move data validation "
              + rawRange
              + " on sheet '"
              + sheet.getSheetName()
              + "' beyond the maximum row "
              + (ExcelRowSpan.MAX_ROW_INDEX + 1));
    }
    return ExcelSheetStructureSupport.formatRange(
        new ExcelRange(firstRow, lastRow, firstColumn, lastColumn));
  }

  private static String shiftedRangeForColumns(
      XSSFSheet sheet,
      String rawRange,
      int firstRow,
      int lastRow,
      int firstColumn,
      int lastColumn) {
    if (lastColumn > ExcelColumnSpan.MAX_COLUMN_INDEX) {
      throw new IllegalArgumentException(
          "INSERT_COLUMNS cannot move data validation "
              + rawRange
              + " on sheet '"
              + sheet.getSheetName()
              + "' beyond the maximum column "
              + SpreadsheetVersion.EXCEL2007.getLastColumnName());
    }
    return ExcelSheetStructureSupport.formatRange(
        new ExcelRange(firstRow, lastRow, firstColumn, lastColumn));
  }

  private static void shiftValidationFormulas(
      CTDataValidation validation,
      XSSFEvaluationWorkbook evaluationWorkbook,
      FormulaShifter formulaShifter,
      int sheetIndex,
      XSSFSheet sheet) {
    String type = validationType(validation);
    if ("list".equals(type)) {
      String formula1 = validation.isSetFormula1() ? validation.getFormula1() : "";
      if (shouldShiftValidationListFormula(formula1)) {
        validation.setFormula1(
            shiftedFormula(
                formula1,
                FormulaType.DATAVALIDATION_LIST,
                evaluationWorkbook,
                formulaShifter,
                sheetIndex,
                sheet));
      }
      return;
    }
    if (validation.isSetFormula1() && !validation.getFormula1().isBlank()) {
      validation.setFormula1(
          shiftedFormula(
              validation.getFormula1(),
              FormulaType.CELL,
              evaluationWorkbook,
              formulaShifter,
              sheetIndex,
              sheet));
    }
    if (validation.isSetFormula2() && !validation.getFormula2().isBlank()) {
      validation.setFormula2(
          shiftedFormula(
              validation.getFormula2(),
              FormulaType.CELL,
              evaluationWorkbook,
              formulaShifter,
              sheetIndex,
              sheet));
    }
  }

  private static String shiftedFormula(
      String formula,
      FormulaType formulaType,
      XSSFEvaluationWorkbook evaluationWorkbook,
      FormulaShifter formulaShifter,
      int sheetIndex,
      XSSFSheet sheet) {
    try {
      Ptg[] tokens = FormulaParser.parse(formula, evaluationWorkbook, formulaType, sheetIndex);
      if (!formulaShifter.adjustFormula(tokens, sheetIndex)) {
        return formula;
      }
      return FormulaRenderer.toFormulaString(evaluationWorkbook, tokens);
    } catch (RuntimeException exception) {
      throw new IllegalArgumentException(
          "Cannot normalize data-validation formula '"
              + formula
              + "' on sheet '"
              + sheet.getSheetName()
              + "' after the structural insert.",
          exception);
    }
  }

  private static String validationType(CTDataValidation validation) {
    return validation.isSetType() ? validation.getType().toString().toLowerCase(Locale.ROOT) : "";
  }

  private static boolean shouldShiftValidationListFormula(String formula1) {
    return !formula1.isBlank() && !isQuotedListLiteral(formula1);
  }

  private static boolean isQuotedListLiteral(String formula) {
    return formula.length() >= 2 && formula.startsWith("\"") && formula.endsWith("\"");
  }
}
