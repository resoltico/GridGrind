package dev.erst.gridgrind.excel;

import java.util.List;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Applies authored data-validation mutations onto one sheet. */
final class ExcelDataValidationAuthoringSupport {
  void setDataValidation(
      XSSFSheet sheet, String range, ExcelDataValidationDefinition validationDefinition) {
    ExcelRange targetRange = ExcelRange.parse(range);
    ExcelDataValidationRangeSupport.normalizeOverlappingSqref(sheet, List.of(targetRange));

    DataValidationHelper helper = sheet.getDataValidationHelper();
    DataValidationConstraint constraint = toConstraint(helper, validationDefinition.rule());
    CellRangeAddressList regions =
        new CellRangeAddressList(
            targetRange.firstRow(),
            targetRange.lastRow(),
            targetRange.firstColumn(),
            targetRange.lastColumn());
    DataValidation validation = helper.createValidation(constraint, regions);
    validation.setEmptyCellAllowed(validationDefinition.allowBlank());
    validation.setSuppressDropDownArrow(validationDefinition.suppressDropDownArrow());
    if (validationDefinition.prompt().isEmpty()) {
      validation.setShowPromptBox(false);
    } else {
      ExcelDataValidationPrompt prompt = validationDefinition.prompt().orElseThrow();
      validation.setShowPromptBox(prompt.showPromptBox());
      validation.createPromptBox(prompt.title(), prompt.text());
    }
    if (validationDefinition.errorAlert().isEmpty()) {
      validation.setShowErrorBox(false);
    } else {
      ExcelDataValidationErrorAlert errorAlert = validationDefinition.errorAlert().orElseThrow();
      validation.setShowErrorBox(errorAlert.showErrorBox());
      validation.setErrorStyle(ExcelDataValidationPoiBridge.toPoiErrorStyle(errorAlert.style()));
      validation.createErrorBox(errorAlert.title(), errorAlert.text());
    }
    sheet.addValidationData(validation);
    ExcelDataValidationRangeSupport.syncValidationCount(sheet);
  }

  void clearDataValidations(XSSFSheet sheet, ExcelRangeSelection selection) {
    switch (selection) {
      case ExcelRangeSelection.All _ -> {
        if (sheet.getCTWorksheet().isSetDataValidations()) {
          sheet.getCTWorksheet().unsetDataValidations();
        }
      }
      case ExcelRangeSelection.Selected selected ->
          ExcelDataValidationRangeSupport.normalizeOverlappingSqref(
              sheet, selected.ranges().stream().map(ExcelRange::parse).toList());
    }
  }

  static DataValidationConstraint toConstraint(
      DataValidationHelper helper, ExcelDataValidationRule rule) {
    return switch (rule) {
      case ExcelDataValidationRule.ExplicitList explicitList ->
          explicitList.values().isEmpty()
              ? helper.createFormulaListConstraint("\"\"")
              : helper.createExplicitListConstraint(explicitList.values().toArray(String[]::new));
      case ExcelDataValidationRule.FormulaList formulaList ->
          helper.createFormulaListConstraint(formulaList.formula());
      case ExcelDataValidationRule.WholeNumber wholeNumber ->
          helper.createIntegerConstraint(
              ExcelComparisonOperatorPoiBridge.toPoi(wholeNumber.operator()),
              wholeNumber.formula1(),
              wholeNumber.formula2().orElse(null));
      case ExcelDataValidationRule.DecimalNumber decimalNumber ->
          helper.createDecimalConstraint(
              ExcelComparisonOperatorPoiBridge.toPoi(decimalNumber.operator()),
              decimalNumber.formula1(),
              decimalNumber.formula2().orElse(null));
      case ExcelDataValidationRule.DateRule dateRule ->
          helper.createDateConstraint(
              ExcelComparisonOperatorPoiBridge.toPoi(dateRule.operator()),
              dateRule.formula1(),
              dateRule.formula2().orElse(null),
              null);
      case ExcelDataValidationRule.TimeRule timeRule ->
          helper.createTimeConstraint(
              ExcelComparisonOperatorPoiBridge.toPoi(timeRule.operator()),
              timeRule.formula1(),
              timeRule.formula2().orElse(null));
      case ExcelDataValidationRule.TextLength textLength ->
          helper.createTextLengthConstraint(
              ExcelComparisonOperatorPoiBridge.toPoi(textLength.operator()),
              textLength.formula1(),
              textLength.formula2().orElse(null));
      case ExcelDataValidationRule.CustomFormula customFormula ->
          helper.createCustomConstraint(customFormula.formula());
    };
  }
}
