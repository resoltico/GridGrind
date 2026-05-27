package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.validation.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationSnapshot;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Sheet metadata operations such as validations, conditional formats, and autofilters. */
public final class ExcelSheetMetadata {
  private final ExcelSheet sheet;
  private final ExcelSheetMetadataSupport metadataSupport;

  ExcelSheetMetadata(ExcelSheet sheet, ExcelSheetMetadataSupport metadataSupport) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.metadataSupport =
        Objects.requireNonNull(metadataSupport, "metadataSupport must not be null");
  }

  /** Creates or replaces one data-validation rule over the requested sheet range. */
  public ExcelSheetMetadata setDataValidation(
      String range, ExcelDataValidationDefinition validation) {
    metadataSupport.setDataValidation(range, validation, sheet);
    return this;
  }

  /** Removes data-validation structures on the sheet that match the provided range selection. */
  public ExcelSheetMetadata clearDataValidations(ExcelRangeSelection selection) {
    metadataSupport.clearDataValidations(selection, sheet);
    return this;
  }

  /** Creates or replaces one logical conditional-formatting block on this sheet. */
  public ExcelSheetMetadata setConditionalFormatting(
      ExcelConditionalFormattingBlockDefinition block) {
    metadataSupport.setConditionalFormatting(block, sheet);
    return this;
  }

  /** Removes conditional-formatting blocks on this sheet matching the provided selection. */
  public ExcelSheetMetadata clearConditionalFormatting(ExcelRangeSelection selection) {
    metadataSupport.clearConditionalFormatting(selection, sheet);
    return this;
  }

  /** Creates or replaces one sheet-level autofilter range. */
  public ExcelSheetMetadata setAutofilter(String range) {
    metadataSupport.setAutofilter(range, sheet);
    return this;
  }

  /** Creates or replaces one sheet-level autofilter range plus authored criteria and sort state. */
  public ExcelSheetMetadata setAutofilter(
      String range,
      List<ExcelAutofilterFilterColumn> criteria,
      Optional<ExcelAutofilterSortState> sortState) {
    metadataSupport.setAutofilter(range, criteria, sortState, sheet);
    return this;
  }

  /** Clears the sheet-level autofilter range on this sheet. */
  public ExcelSheetMetadata clearAutofilter() {
    metadataSupport.clearAutofilter(sheet);
    return this;
  }

  /** Returns data-validation metadata for the selected ranges on this sheet. */
  public List<ExcelDataValidationSnapshot> dataValidations(ExcelRangeSelection selection) {
    return metadataSupport.dataValidations(selection);
  }

  /** Returns factual conditional-formatting blocks for the selected ranges on this sheet. */
  public List<ExcelConditionalFormattingBlockSnapshot> conditionalFormatting(
      ExcelRangeSelection selection) {
    return metadataSupport.conditionalFormatting(selection);
  }
}
