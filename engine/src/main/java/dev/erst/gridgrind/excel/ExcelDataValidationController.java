package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import java.util.List;
import java.util.Objects;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;

/** Reads, writes, and analyzes Excel data-validation structures on one XSSF sheet. */
final class ExcelDataValidationController {
  private final ExcelDataValidationAuthoringSupport authoring =
      new ExcelDataValidationAuthoringSupport();
  private final ExcelDataValidationSnapshotSupport snapshots =
      new ExcelDataValidationSnapshotSupport();
  private final ExcelDataValidationHealthSupport health = new ExcelDataValidationHealthSupport();

  /** Creates or replaces one data-validation rule over the requested sheet range. */
  void setDataValidation(
      XSSFSheet sheet, String range, ExcelDataValidationDefinition validationDefinition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(range, "range must not be null");
    Objects.requireNonNull(validationDefinition, "validationDefinition must not be null");
    authoring.setDataValidation(sheet, range, validationDefinition);
  }

  /** Removes data-validation structures on the sheet that match the provided range selection. */
  void clearDataValidations(XSSFSheet sheet, ExcelRangeSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    authoring.clearDataValidations(sheet, selection);
  }

  /** Returns data-validation metadata for the selected ranges on one sheet. */
  List<ExcelDataValidationSnapshot> dataValidations(
      XSSFSheet sheet, ExcelRangeSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return snapshots.dataValidations(sheet, selection);
  }

  /** Returns the number of raw data-validation structures currently present on the sheet. */
  int dataValidationCount(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    return dataValidations == null ? 0 : dataValidations.sizeOfDataValidationArray();
  }

  /** Returns derived findings for data-validation health on this sheet. */
  List<WorkbookAnalysis.AnalysisFinding> dataValidationHealthFindings(
      String sheetName, XSSFSheet sheet) {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(sheet, "sheet must not be null");
    return health.dataValidationHealthFindings(
        sheetName, dataValidations(sheet, new ExcelRangeSelection.All()));
  }

  static ExcelDataValidationSnapshot toSnapshot(
      XSSFDataValidation validation, List<String> ranges) {
    return ExcelDataValidationSnapshotSupport.toSnapshot(validation, ranges);
  }

  static ExcelDataValidationSnapshot toSnapshot(CTDataValidation validation, List<String> ranges) {
    return ExcelDataValidationSnapshotSupport.toSnapshot(validation, ranges);
  }

  static ExcelDataValidationPrompt prompt(XSSFDataValidation validation) {
    return ExcelDataValidationSnapshotSupport.prompt(validation);
  }

  static ExcelDataValidationPrompt prompt(CTDataValidation validation) {
    return ExcelDataValidationSnapshotSupport.prompt(validation);
  }

  static ExcelDataValidationErrorAlert errorAlert(XSSFDataValidation validation) {
    return ExcelDataValidationSnapshotSupport.errorAlert(validation);
  }

  static ExcelDataValidationErrorAlert errorAlert(CTDataValidation validation) {
    return ExcelDataValidationSnapshotSupport.errorAlert(validation);
  }

  static ExcelComparisonOperator comparisonOperator(CTDataValidation validation) {
    return ExcelDataValidationSnapshotSupport.comparisonOperator(validation);
  }

  static ExcelDataValidationErrorStyle errorStyle(CTDataValidation validation) {
    return ExcelDataValidationSnapshotSupport.errorStyle(validation);
  }

  static String operatorLabel(ExcelComparisonOperator operator) {
    return ExcelDataValidationSnapshotSupport.operatorLabel(operator);
  }
}
