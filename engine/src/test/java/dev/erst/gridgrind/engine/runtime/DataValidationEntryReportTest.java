package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.excel.ExcelComparisonFormulaSupport;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for protocol-facing data-validation read reports. */
class DataValidationEntryReportTest {
  @Test
  void supportedAndUnsupportedEntriesCopyRangesAndValidateFields() {
    List<RangeLabel> ranges = new ArrayList<>(List.of(new RangeLabel("A1:A3")));

    DataValidationEntryReport.Supported supported =
        new DataValidationEntryReport.Supported(
            ranges,
            new DataValidationEntryReport.DataValidationDefinitionReport(
                new DataValidationRuleSpec.ExplicitList(List.of("Queued", "Done")),
                true,
                false,
                Optional.empty(),
                Optional.empty()));
    DataValidationEntryReport.Unsupported unsupported =
        new DataValidationEntryReport.Unsupported(
            List.of(new RangeLabel("B1:B3")), "ANY", "Not modeled");
    ranges.clear();

    assertEquals(
        List.of("A1:A3"),
        supported.ranges().stream()
            .map(dev.erst.gridgrind.core.vocabulary.WireString::value)
            .toList());
    assertEquals("ANY", unsupported.kind());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationEntryReport.Supported(
                List.of(),
                new DataValidationEntryReport.DataValidationDefinitionReport(
                    new DataValidationRuleSpec.ExplicitList(List.of("Queued")),
                    false,
                    false,
                    Optional.empty(),
                    Optional.empty())));
    assertDoesNotThrow(
        () ->
            new DataValidationEntryReport.Unsupported(
                List.of(new RangeLabel(" ")), "ANY", "detail"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationEntryReport.Unsupported(
                List.of(new RangeLabel("A1")), " ", "detail"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationEntryReport.Unsupported(List.of(new RangeLabel("A1")), "ANY", " "));
  }

  @Test
  void fromExcelCoversEveryRuleFamilyAndNestedMetadata() {
    DataValidationEntryReport.DataValidationDefinitionReport explicitList =
        reportFor(
            new ResolvedDataValidation(
                new DataValidationRuleSpec.ExplicitList(List.of("Queued", "Done")),
                true,
                false,
                Optional.of(
                    new ResolvedValidationPrompt("Status", "Choose one workflow state.", true)),
                Optional.of(
                    new ResolvedValidationErrorAlert(
                        ExcelDataValidationErrorStyle.STOP,
                        "Invalid status",
                        "Use one of the allowed values.",
                        true))));
    DataValidationEntryReport.DataValidationDefinitionReport formulaList =
        reportFor(
            new ResolvedDataValidation(
                new DataValidationRuleSpec.FormulaList(
                    ExcelComparisonFormulaSupport.normalizeFormula("Statuses", "formula")),
                false,
                false,
                Optional.empty(),
                Optional.empty()));
    DataValidationEntryReport.DataValidationDefinitionReport wholeNumber =
        reportFor(
            new ResolvedDataValidation(
                new DataValidationRuleSpec.WholeNumber(
                    ExcelComparisonOperator.GREATER_THAN,
                    ExcelComparisonFormulaSupport.normalizeFormula("1", "formula1"),
                    ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                        Optional.empty(), "formula2")),
                false,
                false,
                Optional.empty(),
                Optional.empty()));
    DataValidationEntryReport.DataValidationDefinitionReport decimal =
        reportFor(
            new ResolvedDataValidation(
                new DataValidationRuleSpec.DecimalNumber(
                    ExcelComparisonOperator.GREATER_THAN,
                    ExcelComparisonFormulaSupport.normalizeFormula("0.5", "formula1"),
                    ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                        Optional.empty(), "formula2")),
                false,
                false,
                Optional.empty(),
                Optional.empty()));
    DataValidationEntryReport.DataValidationDefinitionReport date =
        reportFor(
            new ResolvedDataValidation(
                new DataValidationRuleSpec.DateRule(
                    ExcelComparisonOperator.GREATER_OR_EQUAL,
                    ExcelComparisonFormulaSupport.normalizeFormula("TODAY()", "formula1"),
                    ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                        Optional.empty(), "formula2")),
                false,
                false,
                Optional.empty(),
                Optional.empty()));
    DataValidationEntryReport.DataValidationDefinitionReport time =
        reportFor(
            new ResolvedDataValidation(
                new DataValidationRuleSpec.TimeRule(
                    ExcelComparisonOperator.GREATER_THAN,
                    ExcelComparisonFormulaSupport.normalizeFormula("TIME(9,0,0)", "formula1"),
                    ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                        Optional.empty(), "formula2")),
                false,
                false,
                Optional.empty(),
                Optional.empty()));
    DataValidationEntryReport.DataValidationDefinitionReport textLength =
        reportFor(
            new ResolvedDataValidation(
                new DataValidationRuleSpec.TextLength(
                    ExcelComparisonOperator.LESS_OR_EQUAL,
                    ExcelComparisonFormulaSupport.normalizeFormula("20", "formula1"),
                    ExcelComparisonFormulaSupport.normalizeOptionalFormula(
                        Optional.empty(), "formula2")),
                false,
                false,
                Optional.empty(),
                Optional.empty()));
    DataValidationEntryReport.DataValidationDefinitionReport custom =
        reportFor(
            new ResolvedDataValidation(
                new DataValidationRuleSpec.CustomFormula(
                    ExcelComparisonFormulaSupport.normalizeFormula("LEN(A1)>0", "formula")),
                false,
                false,
                Optional.empty(),
                Optional.empty()));

    assertInstanceOf(DataValidationRuleSpec.ExplicitList.class, explicitList.rule());
    assertTrue(explicitList.prompt().isPresent());
    assertTrue(explicitList.errorAlert().isPresent());
    assertInstanceOf(DataValidationRuleSpec.FormulaList.class, formulaList.rule());
    assertInstanceOf(DataValidationRuleSpec.WholeNumber.class, wholeNumber.rule());
    assertInstanceOf(DataValidationRuleSpec.DecimalNumber.class, decimal.rule());
    assertInstanceOf(DataValidationRuleSpec.DateRule.class, date.rule());
    assertInstanceOf(DataValidationRuleSpec.TimeRule.class, time.rule());
    assertInstanceOf(DataValidationRuleSpec.TextLength.class, textLength.rule());
    assertInstanceOf(DataValidationRuleSpec.CustomFormula.class, custom.rule());
    assertThrows(
        NullPointerException.class,
        () -> InspectionResultValidationReportSupport.toDataValidationDefinitionReport(null));
  }

  @Test
  void fromExcelPreservesEmptyExplicitLists() {
    DataValidationEntryReport.DataValidationDefinitionReport report =
        reportFor(
            new ResolvedDataValidation(
                new DataValidationRuleSpec.ExplicitList(List.of()),
                false,
                false,
                Optional.empty(),
                Optional.empty()));

    DataValidationRuleSpec.ExplicitList rule =
        assertInstanceOf(DataValidationRuleSpec.ExplicitList.class, report.rule());
    assertEquals(List.of(), rule.values());
  }

  private static DataValidationEntryReport.DataValidationDefinitionReport reportFor(
      ResolvedDataValidation definition) {
    return InspectionResultValidationReportSupport.toDataValidationDefinitionReport(definition);
  }
}
