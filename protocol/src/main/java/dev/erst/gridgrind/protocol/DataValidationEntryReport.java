package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import java.util.List;
import java.util.Objects;

/** Protocol-facing factual report for one data-validation structure read from a sheet. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = DataValidationEntryReport.Supported.class, name = "SUPPORTED"),
  @JsonSubTypes.Type(value = DataValidationEntryReport.Unsupported.class, name = "UNSUPPORTED")
})
public sealed interface DataValidationEntryReport
    permits DataValidationEntryReport.Supported, DataValidationEntryReport.Unsupported {

  /** Covered A1-style ranges stored on the sheet for this validation structure. */
  List<String> ranges();

  /** Fully supported data-validation structure. */
  record Supported(List<String> ranges, DataValidationDefinitionReport validation)
      implements DataValidationEntryReport {
    public Supported {
      ranges = copyRanges(ranges);
      Objects.requireNonNull(validation, "validation must not be null");
    }
  }

  /** Present workbook structure that GridGrind can detect but not yet model fully. */
  record Unsupported(List<String> ranges, String kind, String detail)
      implements DataValidationEntryReport {
    public Unsupported {
      ranges = copyRanges(ranges);
      kind = requireNonBlank(kind, "kind");
      detail = requireNonBlank(detail, "detail");
    }
  }

  /** Protocol-facing supported validation definition. */
  record DataValidationDefinitionReport(
      DataValidationRuleInput rule,
      boolean allowBlank,
      boolean suppressDropDownArrow,
      DataValidationPromptInput prompt,
      DataValidationErrorAlertInput errorAlert) {
    public DataValidationDefinitionReport {
      Objects.requireNonNull(rule, "rule must not be null");
    }

    static DataValidationDefinitionReport fromExcel(ExcelDataValidationDefinition definition) {
      Objects.requireNonNull(definition, "definition must not be null");
      return new DataValidationDefinitionReport(
          fromExcelRule(definition.rule()),
          definition.allowBlank(),
          definition.suppressDropDownArrow(),
          definition.prompt() == null
              ? null
              : new DataValidationPromptInput(
                  definition.prompt().title(),
                  definition.prompt().text(),
                  definition.prompt().showPromptBox()),
          definition.errorAlert() == null
              ? null
              : new DataValidationErrorAlertInput(
                  definition.errorAlert().style(),
                  definition.errorAlert().title(),
                  definition.errorAlert().text(),
                  definition.errorAlert().showErrorBox()));
    }

    private static DataValidationRuleInput fromExcelRule(ExcelDataValidationRule rule) {
      return switch (rule) {
        case ExcelDataValidationRule.ExplicitList explicitList ->
            new DataValidationRuleInput.ExplicitList(explicitList.values());
        case ExcelDataValidationRule.FormulaList formulaList ->
            new DataValidationRuleInput.FormulaList(formulaList.formula());
        case ExcelDataValidationRule.WholeNumber wholeNumber ->
            new DataValidationRuleInput.WholeNumber(
                wholeNumber.operator(), wholeNumber.formula1(), wholeNumber.formula2());
        case ExcelDataValidationRule.DecimalNumber decimalNumber ->
            new DataValidationRuleInput.DecimalNumber(
                decimalNumber.operator(), decimalNumber.formula1(), decimalNumber.formula2());
        case ExcelDataValidationRule.DateRule dateRule ->
            new DataValidationRuleInput.DateRule(
                dateRule.operator(), dateRule.formula1(), dateRule.formula2());
        case ExcelDataValidationRule.TimeRule timeRule ->
            new DataValidationRuleInput.TimeRule(
                timeRule.operator(), timeRule.formula1(), timeRule.formula2());
        case ExcelDataValidationRule.TextLength textLength ->
            new DataValidationRuleInput.TextLength(
                textLength.operator(), textLength.formula1(), textLength.formula2());
        case ExcelDataValidationRule.CustomFormula customFormula ->
            new DataValidationRuleInput.CustomFormula(customFormula.formula());
      };
    }
  }

  private static List<String> copyRanges(List<String> ranges) {
    Objects.requireNonNull(ranges, "ranges must not be null");
    List<String> copy = List.copyOf(ranges);
    if (copy.isEmpty()) {
      throw new IllegalArgumentException("ranges must not be empty");
    }
    for (String range : copy) {
      requireNonBlank(range, "ranges");
    }
    return copy;
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
