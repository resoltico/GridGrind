package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import org.apache.poi.ss.formula.FormulaType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;

/** Rebuilds copied data-validation XML and retargets workbook formulas to the new sheet. */
final class ExcelSheetCopyValidationSupport {
  private ExcelSheetCopyValidationSupport() {}

  static void replaceDataValidations(
      List<CTDataValidation> validations,
      XSSFSheet targetPoiSheet,
      ExcelWorkbook workbook,
      String sourceSheetName,
      String newSheetName) {
    if (targetPoiSheet.getCTWorksheet().isSetDataValidations()) {
      targetPoiSheet.getCTWorksheet().unsetDataValidations();
    }
    copyDataValidations(validations, targetPoiSheet, workbook, sourceSheetName, newSheetName);
  }

  static List<CTDataValidation> copiedDataValidations(XSSFSheet sourcePoiSheet) {
    if (!sourcePoiSheet.getCTWorksheet().isSetDataValidations()) {
      return List.of();
    }
    List<CTDataValidation> copiedValidations = new ArrayList<>();
    for (CTDataValidation validation :
        sourcePoiSheet.getCTWorksheet().getDataValidations().getDataValidationArray()) {
      copiedValidations.add((CTDataValidation) validation.copy());
    }
    return List.copyOf(copiedValidations);
  }

  private static void copyDataValidations(
      List<CTDataValidation> validations,
      XSSFSheet targetPoiSheet,
      ExcelWorkbook workbook,
      String sourceSheetName,
      String newSheetName) {
    if (validations.isEmpty()) {
      return;
    }
    var targetDataValidations = targetPoiSheet.getCTWorksheet().addNewDataValidations();
    int targetSheetIndex = workbook.xssfWorkbook().getSheetIndex(newSheetName);
    for (CTDataValidation validation : validations) {
      CTDataValidation copiedValidation = targetDataValidations.addNewDataValidation();
      copiedValidation.set(validation);
      retargetValidationFormulas(
          workbook, copiedValidation, targetSheetIndex, sourceSheetName, newSheetName);
    }
    targetDataValidations.setCount(targetDataValidations.sizeOfDataValidationArray());
  }

  private static void retargetValidationFormulas(
      ExcelWorkbook workbook,
      CTDataValidation validation,
      int targetSheetIndex,
      String sourceSheetName,
      String newSheetName) {
    String type = validationType(validation);
    if ("list".equals(type)) {
      String formula1 = validation.isSetFormula1() ? validation.getFormula1() : "";
      if (shouldRetargetValidationListFormula(formula1)) {
        validation.setFormula1(
            ExcelSheetCopySupport.retargetFormula(
                workbook,
                formula1,
                FormulaType.DATAVALIDATION_LIST,
                targetSheetIndex,
                sourceSheetName,
                newSheetName));
      }
      return;
    }
    if (validation.isSetFormula1() && !validation.getFormula1().isBlank()) {
      validation.setFormula1(
          ExcelSheetCopySupport.retargetFormula(
              workbook,
              validation.getFormula1(),
              FormulaType.CELL,
              targetSheetIndex,
              sourceSheetName,
              newSheetName));
    }
    if (validation.isSetFormula2() && !validation.getFormula2().isBlank()) {
      validation.setFormula2(
          ExcelSheetCopySupport.retargetFormula(
              workbook,
              validation.getFormula2(),
              FormulaType.CELL,
              targetSheetIndex,
              sourceSheetName,
              newSheetName));
    }
  }

  private static String validationType(CTDataValidation validation) {
    return validation.isSetType() ? validation.getType().toString().toLowerCase(Locale.ROOT) : "";
  }

  private static boolean shouldRetargetValidationListFormula(String formula1) {
    return !formula1.isBlank() && !isQuotedListLiteral(formula1);
  }

  private static boolean isQuotedListLiteral(String formula) {
    return formula.length() >= 2 && formula.startsWith("\"") && formula.endsWith("\"");
  }
}
