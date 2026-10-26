package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationController;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Sheet-owned validation, conditional-formatting, and autofilter support. */
final class ExcelSheetMetadataSupport {
  private final Sheet sheet;
  private final ExcelDataValidationController dataValidationController;
  private final ExcelConditionalFormattingController conditionalFormattingController;
  private final ExcelAutofilterController autofilterController;

  ExcelSheetMetadataSupport(Sheet sheet) {
    this(
        sheet,
        new ExcelDataValidationController(),
        new ExcelConditionalFormattingController(),
        new ExcelAutofilterController());
  }

  ExcelSheetMetadataSupport(
      Sheet sheet,
      ExcelDataValidationController dataValidationController,
      ExcelConditionalFormattingController conditionalFormattingController,
      ExcelAutofilterController autofilterController) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.dataValidationController =
        Objects.requireNonNull(
            dataValidationController, "dataValidationController must not be null");
    this.conditionalFormattingController =
        Objects.requireNonNull(
            conditionalFormattingController, "conditionalFormattingController must not be null");
    this.autofilterController =
        Objects.requireNonNull(autofilterController, "autofilterController must not be null");
  }

  ExcelSheet setDataValidation(String range, ResolvedDataValidation validation, ExcelSheet owner) {
    requireNonBlank(range, "range");
    Objects.requireNonNull(validation, "validation must not be null");
    dataValidationController.setDataValidation(xssfSheet(), range, validation);
    return owner;
  }

  ExcelSheet clearDataValidations(ExcelRangeSelection selection, ExcelSheet owner) {
    Objects.requireNonNull(selection, "selection must not be null");
    dataValidationController.clearDataValidations(xssfSheet(), selection);
    return owner;
  }

  ExcelSheet setConditionalFormatting(
      List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges,
      dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec formatting,
      ExcelSheet owner) {
    Objects.requireNonNull(formatting, "formatting must not be null");
    conditionalFormattingController.setConditionalFormatting(xssfSheet(), ranges, formatting);
    return owner;
  }

  ExcelSheet clearConditionalFormatting(ExcelRangeSelection selection, ExcelSheet owner) {
    Objects.requireNonNull(selection, "selection must not be null");
    conditionalFormattingController.clearConditionalFormatting(xssfSheet(), selection);
    return owner;
  }

  ExcelSheet setAutofilter(String range, ExcelSheet owner) {
    return setAutofilter(range, List.of(), Optional.empty(), owner);
  }

  ExcelSheet setAutofilter(
      String range,
      List<AutofilterFilterColumnSpec> criteria,
      Optional<AutofilterSortStateSpec> sortState,
      ExcelSheet owner) {
    requireNonBlank(range, "range");
    Objects.requireNonNull(criteria, "criteria must not be null");
    Objects.requireNonNull(sortState, "sortState must not be null");
    autofilterController.setSheetAutofilter(xssfSheet(), range, criteria, sortState);
    return owner;
  }

  ExcelSheet clearAutofilter(ExcelSheet owner) {
    autofilterController.clearSheetAutofilter(xssfSheet());
    return owner;
  }

  List<DataValidationEntryReport> dataValidations(ExcelRangeSelection selection) {
    Objects.requireNonNull(selection, "selection must not be null");
    return dataValidationController.dataValidations(xssfSheet(), selection);
  }

  List<ConditionalFormattingEntryReport> conditionalFormatting(ExcelRangeSelection selection) {
    Objects.requireNonNull(selection, "selection must not be null");
    return conditionalFormattingController.conditionalFormatting(xssfSheet(), selection);
  }

  int conditionalFormattingBlockCount() {
    return conditionalFormattingController.conditionalFormattingBlockCount(xssfSheet());
  }

  List<AnalysisFindingReport> conditionalFormattingHealthFindings(String sheetName) {
    return conditionalFormattingController.conditionalFormattingHealthFindings(
        sheetName, xssfSheet());
  }

  private XSSFSheet xssfSheet() {
    return (XSSFSheet) sheet;
  }

  private static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
  }
}
