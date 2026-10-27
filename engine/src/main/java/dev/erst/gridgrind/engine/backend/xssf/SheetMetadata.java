package dev.erst.gridgrind.engine.backend.xssf;

import dev.erst.gridgrind.core.model.table.AutofilterFilterColumnSpec;
import dev.erst.gridgrind.core.model.table.AutofilterSortStateSpec;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.engine.autofilter.AutofilterAdapter;
import dev.erst.gridgrind.engine.conditionalformatting.ConditionalFormattingAdapter;
import dev.erst.gridgrind.engine.validation.DataValidationAdapter;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Sheet-owned validation, conditional-formatting, and autofilter support. */
public final class SheetMetadata {
  private final Sheet sheet;
  private final DataValidationAdapter dataValidationAdapter;
  private final ConditionalFormattingAdapter conditionalFormattingAdapter;
  private final AutofilterAdapter autofilterAdapter;

  SheetMetadata(Sheet sheet) {
    this(
        sheet,
        new DataValidationAdapter(),
        new ConditionalFormattingAdapter(),
        new AutofilterAdapter());
  }

  SheetMetadata(
      Sheet sheet,
      DataValidationAdapter dataValidationAdapter,
      ConditionalFormattingAdapter conditionalFormattingAdapter,
      AutofilterAdapter autofilterAdapter) {
    this.sheet = Objects.requireNonNull(sheet, "sheet must not be null");
    this.dataValidationAdapter =
        Objects.requireNonNull(dataValidationAdapter, "dataValidationAdapter must not be null");
    this.conditionalFormattingAdapter =
        Objects.requireNonNull(
            conditionalFormattingAdapter, "conditionalFormattingAdapter must not be null");
    this.autofilterAdapter =
        Objects.requireNonNull(autofilterAdapter, "autofilterAdapter must not be null");
  }

  /** Creates or replaces the validation rule over the selected range. */
  public SheetMetadata setDataValidation(String range, ResolvedDataValidation validation) {
    requireNonBlank(range, "range");
    Objects.requireNonNull(validation, "validation must not be null");
    dataValidationAdapter.setDataValidation(xssfSheet(), range, validation);
    return this;
  }

  /** Creates or replaces the conditional-formatting block over the selected ranges. */
  public SheetMetadata setConditionalFormatting(
      List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges,
      dev.erst.gridgrind.core.model.validation.ConditionalFormattingSpec formatting) {
    Objects.requireNonNull(formatting, "formatting must not be null");
    conditionalFormattingAdapter.setConditionalFormatting(xssfSheet(), ranges, formatting);
    return this;
  }

  /** Creates or replaces the sheet autofilter and any authored criteria. */
  public SheetMetadata setAutofilter(String range) {
    return setAutofilter(range, List.of(), Optional.empty());
  }

  /** Creates or replaces the sheet autofilter and any authored criteria. */
  public SheetMetadata setAutofilter(
      String range,
      List<AutofilterFilterColumnSpec> criteria,
      Optional<AutofilterSortStateSpec> sortState) {
    requireNonBlank(range, "range");
    Objects.requireNonNull(criteria, "criteria must not be null");
    Objects.requireNonNull(sortState, "sortState must not be null");
    autofilterAdapter.setSheetAutofilter(xssfSheet(), range, criteria, sortState);
    return this;
  }

  /** Removes the sheet-owned autofilter. */
  public SheetMetadata clearAutofilter() {
    autofilterAdapter.clearSheetAutofilter(xssfSheet());
    return this;
  }

  int conditionalFormattingBlockCount() {
    return conditionalFormattingAdapter.conditionalFormattingBlockCount(xssfSheet());
  }

  List<AnalysisFindingReport> conditionalFormattingHealthFindings(String sheetName) {
    return conditionalFormattingAdapter.conditionalFormattingHealthFindings(sheetName, xssfSheet());
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

  /** Clears data-validation rules on the bound sheet. */
  public SheetMetadata clearDataValidations() {
    dataValidationAdapter.clearDataValidations(xssfSheet());
    return this;
  }

  /** Clears data-validation rules on the bound sheet. */
  public SheetMetadata clearDataValidations(
      List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges) {
    dataValidationAdapter.clearDataValidations(xssfSheet(), ranges);
    return this;
  }

  /** Clears conditional-formatting rules on the bound sheet. */
  public SheetMetadata clearConditionalFormatting() {
    conditionalFormattingAdapter.clearConditionalFormatting(xssfSheet());
    return this;
  }

  /** Clears conditional-formatting rules on the bound sheet. */
  public SheetMetadata clearConditionalFormatting(
      List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges) {
    conditionalFormattingAdapter.clearConditionalFormatting(xssfSheet(), ranges);
    return this;
  }

  /** Reads effective data-validation rules on the bound sheet. */
  public List<DataValidationEntryReport> dataValidations() {
    return dataValidationAdapter.dataValidations(xssfSheet());
  }

  /** Reads effective data-validation rules on the bound sheet. */
  public List<DataValidationEntryReport> dataValidations(
      List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges) {
    return dataValidationAdapter.dataValidations(xssfSheet(), ranges);
  }

  /** Reads effective conditional-formatting blocks on the bound sheet. */
  public List<ConditionalFormattingEntryReport> conditionalFormatting() {
    return conditionalFormattingAdapter.conditionalFormatting(xssfSheet());
  }

  /** Reads effective conditional-formatting blocks on the bound sheet. */
  public List<ConditionalFormattingEntryReport> conditionalFormatting(
      List<dev.erst.gridgrind.core.vocabulary.RangeRef> ranges) {
    return conditionalFormattingAdapter.conditionalFormatting(xssfSheet(), ranges);
  }
}
