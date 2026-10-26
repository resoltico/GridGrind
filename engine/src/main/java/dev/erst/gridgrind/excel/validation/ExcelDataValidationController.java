package dev.erst.gridgrind.excel.validation;

import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationErrorAlertReport;
import dev.erst.gridgrind.core.model.validation.DataValidationPromptReport;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.ExcelRangeSelection;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;

/** Reads, writes, and analyzes Excel data-validation structures on one XSSF sheet. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelDataValidationController {
  private final ExcelDataValidationAuthoringSupport authoring =
      new ExcelDataValidationAuthoringSupport();
  private final ExcelDataValidationSnapshotSupport snapshots =
      new ExcelDataValidationSnapshotSupport();
  private final ExcelDataValidationHealthSupport health = new ExcelDataValidationHealthSupport();

  /** Creates or replaces one data-validation rule over the requested sheet range. */
  public void setDataValidation(
      XSSFSheet sheet, String range, ResolvedDataValidation validationDefinition) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(range, "range must not be null");
    Objects.requireNonNull(validationDefinition, "validationDefinition must not be null");
    authoring.setDataValidation(sheet, range, validationDefinition);
  }

  /** Removes data-validation structures on the sheet that match the provided range selection. */
  public void clearDataValidations(XSSFSheet sheet, ExcelRangeSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    authoring.clearDataValidations(sheet, selection);
  }

  /** Returns data-validation metadata for the selected ranges on one sheet. */
  public List<DataValidationEntryReport> dataValidations(
      XSSFSheet sheet, ExcelRangeSelection selection) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(selection, "selection must not be null");
    return snapshots.dataValidations(sheet, selection);
  }

  /** Returns the number of raw data-validation structures currently present on the sheet. */
  public int dataValidationCount(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    return dataValidations == null ? 0 : dataValidations.sizeOfDataValidationArray();
  }

  /** Returns derived findings for data-validation health on this sheet. */
  public List<AnalysisFindingReport> dataValidationHealthFindings(
      String sheetName, XSSFSheet sheet) {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(sheet, "sheet must not be null");
    return health.dataValidationHealthFindings(
        sheetName, dataValidations(sheet, new ExcelRangeSelection.All()));
  }

  public static DataValidationEntryReport toSnapshot(
      XSSFDataValidation validation, List<String> ranges) {
    return ExcelDataValidationSnapshotSupport.toSnapshot(validation, ranges);
  }

  public static DataValidationEntryReport toSnapshot(
      CTDataValidation validation, List<String> ranges) {
    return ExcelDataValidationSnapshotSupport.toSnapshot(validation, ranges);
  }

  public static Optional<DataValidationPromptReport> prompt(XSSFDataValidation validation) {
    return ExcelDataValidationSnapshotSupport.prompt(validation);
  }

  public static Optional<DataValidationPromptReport> prompt(CTDataValidation validation) {
    return ExcelDataValidationSnapshotSupport.prompt(validation);
  }

  public static Optional<DataValidationErrorAlertReport> errorAlert(XSSFDataValidation validation) {
    return ExcelDataValidationSnapshotSupport.errorAlert(validation);
  }

  public static Optional<DataValidationErrorAlertReport> errorAlert(CTDataValidation validation) {
    return ExcelDataValidationSnapshotSupport.errorAlert(validation);
  }

  public static ExcelComparisonOperator comparisonOperator(CTDataValidation validation) {
    return ExcelDataValidationSnapshotSupport.comparisonOperator(validation);
  }

  public static ExcelDataValidationErrorStyle errorStyle(CTDataValidation validation) {
    return ExcelDataValidationSnapshotSupport.errorStyle(validation);
  }

  public static String operatorLabel(ExcelComparisonOperator operator) {
    return ExcelDataValidationSnapshotSupport.operatorLabel(operator);
  }
}
