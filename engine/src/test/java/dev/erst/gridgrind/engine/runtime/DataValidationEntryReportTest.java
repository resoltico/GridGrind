package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleInput;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.RangeLabel;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationPrompt;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationRule;
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
                new DataValidationRuleInput.ExplicitList(List.of("Queued", "Done")),
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
                    new DataValidationRuleInput.ExplicitList(List.of("Queued")),
                    false,
                    false,
                    Optional.empty(),
                    Optional.empty())));
    assertThrows(
        IllegalArgumentException.class,
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
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done")),
                true,
                false,
                Optional.of(
                    new ExcelDataValidationPrompt("Status", "Choose one workflow state.", true)),
                Optional.of(
                    new ExcelDataValidationErrorAlert(
                        ExcelDataValidationErrorStyle.STOP,
                        "Invalid status",
                        "Use one of the allowed values.",
                        true))));
    DataValidationEntryReport.DataValidationDefinitionReport formulaList =
        reportFor(
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.FormulaList("Statuses"),
                false,
                false,
                Optional.empty(),
                Optional.empty()));
    DataValidationEntryReport.DataValidationDefinitionReport wholeNumber =
        reportFor(
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.WholeNumber(
                    ExcelComparisonOperator.GREATER_THAN, "1", Optional.empty()),
                false,
                false,
                Optional.empty(),
                Optional.empty()));
    DataValidationEntryReport.DataValidationDefinitionReport decimal =
        reportFor(
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.DecimalNumber(
                    ExcelComparisonOperator.GREATER_THAN, "0.5", Optional.empty()),
                false,
                false,
                Optional.empty(),
                Optional.empty()));
    DataValidationEntryReport.DataValidationDefinitionReport date =
        reportFor(
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.DateRule(
                    ExcelComparisonOperator.GREATER_OR_EQUAL, "TODAY()", Optional.empty()),
                false,
                false,
                Optional.empty(),
                Optional.empty()));
    DataValidationEntryReport.DataValidationDefinitionReport time =
        reportFor(
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.TimeRule(
                    ExcelComparisonOperator.GREATER_THAN, "TIME(9,0,0)", Optional.empty()),
                false,
                false,
                Optional.empty(),
                Optional.empty()));
    DataValidationEntryReport.DataValidationDefinitionReport textLength =
        reportFor(
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.TextLength(
                    ExcelComparisonOperator.LESS_OR_EQUAL, "20", Optional.empty()),
                false,
                false,
                Optional.empty(),
                Optional.empty()));
    DataValidationEntryReport.DataValidationDefinitionReport custom =
        reportFor(
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.CustomFormula("LEN(A1)>0"),
                false,
                false,
                Optional.empty(),
                Optional.empty()));

    assertInstanceOf(DataValidationRuleInput.ExplicitList.class, explicitList.rule());
    assertTrue(explicitList.prompt().isPresent());
    assertTrue(explicitList.errorAlert().isPresent());
    assertInstanceOf(DataValidationRuleInput.FormulaList.class, formulaList.rule());
    assertInstanceOf(DataValidationRuleInput.WholeNumber.class, wholeNumber.rule());
    assertInstanceOf(DataValidationRuleInput.DecimalNumber.class, decimal.rule());
    assertInstanceOf(DataValidationRuleInput.DateRule.class, date.rule());
    assertInstanceOf(DataValidationRuleInput.TimeRule.class, time.rule());
    assertInstanceOf(DataValidationRuleInput.TextLength.class, textLength.rule());
    assertInstanceOf(DataValidationRuleInput.CustomFormula.class, custom.rule());
    assertThrows(
        NullPointerException.class,
        () -> InspectionResultValidationReportSupport.toDataValidationDefinitionReport(null));
  }

  @Test
  void fromExcelPreservesEmptyExplicitLists() {
    DataValidationEntryReport.DataValidationDefinitionReport report =
        reportFor(
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.ExplicitList(List.of()),
                false,
                false,
                Optional.empty(),
                Optional.empty()));

    DataValidationRuleInput.ExplicitList rule =
        assertInstanceOf(DataValidationRuleInput.ExplicitList.class, report.rule());
    assertEquals(List.of(), rule.values());
  }

  private static DataValidationEntryReport.DataValidationDefinitionReport reportFor(
      ExcelDataValidationDefinition definition) {
    return InspectionResultValidationReportSupport.toDataValidationDefinitionReport(definition);
  }
}
