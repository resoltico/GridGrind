package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.ExcelDataValidationPrompt;
import dev.erst.gridgrind.excel.ExcelDataValidationRule;
import dev.erst.gridgrind.protocol.dto.DataValidationEntryReport;
import dev.erst.gridgrind.protocol.dto.DataValidationRuleInput;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for protocol-facing data-validation read reports. */
class DataValidationEntryReportTest {
  @Test
  void supportedAndUnsupportedEntriesCopyRangesAndValidateFields() {
    List<String> ranges = new ArrayList<>(List.of("A1:A3"));

    DataValidationEntryReport.Supported supported =
        new DataValidationEntryReport.Supported(
            ranges,
            new DataValidationEntryReport.DataValidationDefinitionReport(
                new DataValidationRuleInput.ExplicitList(List.of("Queued", "Done")),
                true,
                false,
                null,
                null));
    DataValidationEntryReport.Unsupported unsupported =
        new DataValidationEntryReport.Unsupported(List.of("B1:B3"), "ANY", "Not modeled");
    ranges.clear();

    assertEquals(List.of("A1:A3"), supported.ranges());
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
                    null,
                    null)));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationEntryReport.Unsupported(List.of(" "), "ANY", "detail"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationEntryReport.Unsupported(List.of("A1"), " ", "detail"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationEntryReport.Unsupported(List.of("A1"), "ANY", " "));
  }

  @Test
  void fromExcelCoversEveryRuleFamilyAndNestedMetadata() {
    DataValidationEntryReport.DataValidationDefinitionReport explicitList =
        reportFor(
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done")),
                true,
                false,
                new ExcelDataValidationPrompt("Status", "Choose one workflow state.", true),
                new ExcelDataValidationErrorAlert(
                    ExcelDataValidationErrorStyle.STOP,
                    "Invalid status",
                    "Use one of the allowed values.",
                    true)));
    DataValidationEntryReport.DataValidationDefinitionReport formulaList =
        reportFor(
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.FormulaList("Statuses"), false, false, null, null));
    DataValidationEntryReport.DataValidationDefinitionReport wholeNumber =
        reportFor(
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.WholeNumber(
                    ExcelComparisonOperator.GREATER_THAN, "1", null),
                false,
                false,
                null,
                null));
    DataValidationEntryReport.DataValidationDefinitionReport decimal =
        reportFor(
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.DecimalNumber(
                    ExcelComparisonOperator.GREATER_THAN, "0.5", null),
                false,
                false,
                null,
                null));
    DataValidationEntryReport.DataValidationDefinitionReport date =
        reportFor(
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.DateRule(
                    ExcelComparisonOperator.GREATER_OR_EQUAL, "TODAY()", null),
                false,
                false,
                null,
                null));
    DataValidationEntryReport.DataValidationDefinitionReport time =
        reportFor(
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.TimeRule(
                    ExcelComparisonOperator.GREATER_THAN, "TIME(9,0,0)", null),
                false,
                false,
                null,
                null));
    DataValidationEntryReport.DataValidationDefinitionReport textLength =
        reportFor(
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.TextLength(
                    ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
                false,
                false,
                null,
                null));
    DataValidationEntryReport.DataValidationDefinitionReport custom =
        reportFor(
            new ExcelDataValidationDefinition(
                new ExcelDataValidationRule.CustomFormula("LEN(A1)>0"), false, false, null, null));

    assertInstanceOf(DataValidationRuleInput.ExplicitList.class, explicitList.rule());
    assertNotNull(explicitList.prompt());
    assertNotNull(explicitList.errorAlert());
    assertInstanceOf(DataValidationRuleInput.FormulaList.class, formulaList.rule());
    assertInstanceOf(DataValidationRuleInput.WholeNumber.class, wholeNumber.rule());
    assertInstanceOf(DataValidationRuleInput.DecimalNumber.class, decimal.rule());
    assertInstanceOf(DataValidationRuleInput.DateRule.class, date.rule());
    assertInstanceOf(DataValidationRuleInput.TimeRule.class, time.rule());
    assertInstanceOf(DataValidationRuleInput.TextLength.class, textLength.rule());
    assertInstanceOf(DataValidationRuleInput.CustomFormula.class, custom.rule());
    assertThrows(
        NullPointerException.class,
        () -> WorkbookReadResultConverter.toDataValidationDefinitionReport(null));
  }

  private static DataValidationEntryReport.DataValidationDefinitionReport reportFor(
      ExcelDataValidationDefinition definition) {
    return WorkbookReadResultConverter.toDataValidationDefinitionReport(definition);
  }
}
