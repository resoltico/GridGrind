package dev.erst.gridgrind.excel.validation;

import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.excel.ExcelComparisonFormulaSupport;
import dev.erst.gridgrind.excel.ExcelRange;
import dev.erst.gridgrind.excel.ExcelRangeSelection;
import java.util.List;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Applies authored data-validation mutations onto one sheet. */
public final class ExcelDataValidationAuthoringSupport {
  void setDataValidation(
      XSSFSheet sheet, String range, ResolvedDataValidation validationDefinition) {
    ExcelRange targetRange = ExcelRange.parse(range);
    DataValidationHelper helper = sheet.getDataValidationHelper();
    DataValidationConstraint constraint = toConstraint(helper, validationDefinition.rule());
    ExcelDataValidationRangeSupport.normalizeOverlappingSqref(sheet, List.of(targetRange));
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
      ResolvedValidationPrompt prompt = validationDefinition.prompt().orElseThrow();
      validation.setShowPromptBox(prompt.showPromptBox());
      validation.createPromptBox(prompt.title(), prompt.text());
    }
    if (validationDefinition.errorAlert().isEmpty()) {
      validation.setShowErrorBox(false);
    } else {
      ResolvedValidationErrorAlert errorAlert = validationDefinition.errorAlert().orElseThrow();
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
      DataValidationHelper helper, DataValidationRuleSpec rule) {
    return switch (rule) {
      case DataValidationRuleSpec.ExplicitList explicitList ->
          explicitList.values().isEmpty()
              ? helper.createFormulaListConstraint("\"\"")
              : helper.createExplicitListConstraint(explicitList.values().toArray(String[]::new));
      case DataValidationRuleSpec.FormulaList formulaList ->
          helper.createFormulaListConstraint(
              ExcelComparisonFormulaSupport.normalizeFormula(formulaList.formula(), "formula"));
      case DataValidationRuleSpec.WholeNumber wholeNumber ->
          helper.createIntegerConstraint(
              ExcelDataValidationComparisonOperatorPoiBridge.toPoi(wholeNumber.operator()),
              ExcelComparisonFormulaSupport.normalizeFormula(wholeNumber.formula1(), "formula1"),
              ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                      wholeNumber.formula2(), "formula2")
                  .orElse(null));
      case DataValidationRuleSpec.DecimalNumber decimalNumber ->
          helper.createDecimalConstraint(
              ExcelDataValidationComparisonOperatorPoiBridge.toPoi(decimalNumber.operator()),
              ExcelComparisonFormulaSupport.normalizeFormula(decimalNumber.formula1(), "formula1"),
              ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                      decimalNumber.formula2(), "formula2")
                  .orElse(null));
      case DataValidationRuleSpec.DateRule dateRule ->
          helper.createDateConstraint(
              ExcelDataValidationComparisonOperatorPoiBridge.toPoi(dateRule.operator()),
              ExcelComparisonFormulaSupport.normalizeFormula(dateRule.formula1(), "formula1"),
              ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                      dateRule.formula2(), "formula2")
                  .orElse(null),
              null);
      case DataValidationRuleSpec.TimeRule timeRule ->
          helper.createTimeConstraint(
              ExcelDataValidationComparisonOperatorPoiBridge.toPoi(timeRule.operator()),
              ExcelComparisonFormulaSupport.normalizeFormula(timeRule.formula1(), "formula1"),
              ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                      timeRule.formula2(), "formula2")
                  .orElse(null));
      case DataValidationRuleSpec.TextLength textLength ->
          helper.createTextLengthConstraint(
              ExcelDataValidationComparisonOperatorPoiBridge.toPoi(textLength.operator()),
              ExcelComparisonFormulaSupport.normalizeFormula(textLength.formula1(), "formula1"),
              ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                      textLength.formula2(), "formula2")
                  .orElse(null));
      case DataValidationRuleSpec.CustomFormula customFormula ->
          helper.createCustomConstraint(
              ExcelComparisonFormulaSupport.normalizeFormula(customFormula.formula(), "formula"));
    };
  }
}
