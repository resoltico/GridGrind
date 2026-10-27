package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
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
  public ExcelSheetMetadata setDataValidation(String range, ResolvedDataValidation validation) {
    metadataSupport.setDataValidation(range, validation, sheet);
    return this;
  }

  /** Creates or replaces one logical conditional-formatting block on this sheet. */
  public ExcelSheetMetadata setConditionalFormatting(
      List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges,
      dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec formatting) {
    metadataSupport.setConditionalFormatting(ranges, formatting, sheet);
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
      List<AutofilterFilterColumnSpec> criteria,
      Optional<AutofilterSortStateSpec> sortState) {
    metadataSupport.setAutofilter(range, criteria, sortState, sheet);
    return this;
  }

  /** Clears the sheet-level autofilter range on this sheet. */
  public ExcelSheetMetadata clearAutofilter() {
    metadataSupport.clearAutofilter(sheet);
    return this;
  }

  /** Clears matching metadata on this bound sheet. */
  public ExcelSheetMetadata clearDataValidations() {
    metadataSupport.clearDataValidations(sheet);
    return this;
  }

  /** Clears matching metadata on this bound sheet. */
  public ExcelSheetMetadata clearDataValidations(
      List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges) {
    metadataSupport.clearDataValidations(ranges, sheet);
    return this;
  }

  /** Clears matching metadata on this bound sheet. */
  public ExcelSheetMetadata clearConditionalFormatting() {
    metadataSupport.clearConditionalFormatting(sheet);
    return this;
  }

  /** Clears matching metadata on this bound sheet. */
  public ExcelSheetMetadata clearConditionalFormatting(
      List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges) {
    metadataSupport.clearConditionalFormatting(ranges, sheet);
    return this;
  }

  /** Reads matching metadata on this bound sheet. */
  public List<DataValidationEntryReport> dataValidations() {
    return metadataSupport.dataValidations();
  }

  /** Reads matching metadata on this bound sheet. */
  public List<DataValidationEntryReport> dataValidations(
      List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges) {
    return metadataSupport.dataValidations(ranges);
  }

  /** Reads matching metadata on this bound sheet. */
  public List<ConditionalFormattingEntryReport> conditionalFormatting() {
    return metadataSupport.conditionalFormatting();
  }

  /** Reads matching metadata on this bound sheet. */
  public List<ConditionalFormattingEntryReport> conditionalFormatting(
      List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges) {
    return metadataSupport.conditionalFormatting(ranges);
  }
}
