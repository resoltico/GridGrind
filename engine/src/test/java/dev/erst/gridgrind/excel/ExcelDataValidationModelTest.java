package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationDefinition;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationErrorAlert;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationPoiBridge;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationPrompt;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationRule;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationSnapshot;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.DataValidation;
import org.junit.jupiter.api.Test;

/** Tests for data-validation enums and immutable model records. */
class ExcelDataValidationModelTest {
  @Test
  void comparisonOperatorRoundTripsPoiConstants() {
    assertEquals(
        ExcelComparisonOperator.BETWEEN,
        ExcelConditionalFormattingComparisonOperatorPoiBridge.fromPoi(ComparisonOperator.BETWEEN));
    assertEquals(
        ExcelComparisonOperator.NOT_BETWEEN,
        ExcelConditionalFormattingComparisonOperatorPoiBridge.fromPoi(
            ComparisonOperator.NOT_BETWEEN));
    assertEquals(
        ExcelComparisonOperator.EQUAL,
        ExcelConditionalFormattingComparisonOperatorPoiBridge.fromPoi(ComparisonOperator.EQUAL));
    assertEquals(
        ExcelComparisonOperator.NOT_EQUAL,
        ExcelConditionalFormattingComparisonOperatorPoiBridge.fromPoi(
            ComparisonOperator.NOT_EQUAL));
    assertEquals(
        ExcelComparisonOperator.GREATER_THAN,
        ExcelConditionalFormattingComparisonOperatorPoiBridge.fromPoi(ComparisonOperator.GT));
    assertEquals(
        ExcelComparisonOperator.LESS_THAN,
        ExcelConditionalFormattingComparisonOperatorPoiBridge.fromPoi(ComparisonOperator.LT));
    assertEquals(
        ExcelComparisonOperator.GREATER_OR_EQUAL,
        ExcelConditionalFormattingComparisonOperatorPoiBridge.fromPoi(ComparisonOperator.GE));
    assertEquals(
        ExcelComparisonOperator.LESS_OR_EQUAL,
        ExcelConditionalFormattingComparisonOperatorPoiBridge.fromPoi(ComparisonOperator.LE));
    assertEquals(
        ComparisonOperator.BETWEEN,
        ExcelConditionalFormattingComparisonOperatorPoiBridge.toPoi(
            ExcelComparisonOperator.BETWEEN));
    assertEquals(
        ComparisonOperator.NOT_BETWEEN,
        ExcelConditionalFormattingComparisonOperatorPoiBridge.toPoi(
            ExcelComparisonOperator.NOT_BETWEEN));
    assertEquals(
        ComparisonOperator.EQUAL,
        ExcelConditionalFormattingComparisonOperatorPoiBridge.toPoi(ExcelComparisonOperator.EQUAL));
    assertEquals(
        ComparisonOperator.NOT_EQUAL,
        ExcelConditionalFormattingComparisonOperatorPoiBridge.toPoi(
            ExcelComparisonOperator.NOT_EQUAL));
    assertEquals(
        ComparisonOperator.GT,
        ExcelConditionalFormattingComparisonOperatorPoiBridge.toPoi(
            ExcelComparisonOperator.GREATER_THAN));
    assertEquals(
        ComparisonOperator.LT,
        ExcelConditionalFormattingComparisonOperatorPoiBridge.toPoi(
            ExcelComparisonOperator.LESS_THAN));
    assertEquals(
        ComparisonOperator.GE,
        ExcelConditionalFormattingComparisonOperatorPoiBridge.toPoi(
            ExcelComparisonOperator.GREATER_OR_EQUAL));
    assertEquals(
        ComparisonOperator.LE,
        ExcelConditionalFormattingComparisonOperatorPoiBridge.toPoi(
            ExcelComparisonOperator.LESS_OR_EQUAL));
  }

  @Test
  void comparisonOperatorRejectsUnsupportedPoiValues() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelConditionalFormattingComparisonOperatorPoiBridge.fromPoi(-99));

    assertEquals(
        "Unsupported Apache POI conditional-formatting comparison operator: -99",
        failure.getMessage());
  }

  @Test
  void errorStyleRoundTripsPoiConstants() {
    assertEquals(
        ExcelDataValidationErrorStyle.STOP,
        ExcelDataValidationPoiBridge.fromPoiErrorStyle(DataValidation.ErrorStyle.STOP));
    assertEquals(
        ExcelDataValidationErrorStyle.WARNING,
        ExcelDataValidationPoiBridge.fromPoiErrorStyle(DataValidation.ErrorStyle.WARNING));
    assertEquals(
        ExcelDataValidationErrorStyle.INFORMATION,
        ExcelDataValidationPoiBridge.fromPoiErrorStyle(DataValidation.ErrorStyle.INFO));
    assertEquals(
        DataValidation.ErrorStyle.STOP,
        ExcelDataValidationPoiBridge.toPoiErrorStyle(ExcelDataValidationErrorStyle.STOP));
    assertEquals(
        DataValidation.ErrorStyle.WARNING,
        ExcelDataValidationPoiBridge.toPoiErrorStyle(ExcelDataValidationErrorStyle.WARNING));
    assertEquals(
        DataValidation.ErrorStyle.INFO,
        ExcelDataValidationPoiBridge.toPoiErrorStyle(ExcelDataValidationErrorStyle.INFORMATION));
  }

  @Test
  void errorStyleRejectsUnsupportedPoiValues() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> ExcelDataValidationPoiBridge.fromPoiErrorStyle(-77));

    assertEquals("Unsupported Apache POI error style: -77", failure.getMessage());
  }

  @Test
  void promptAndAlertRejectBlankValues() {
    assertThrows(
        NullPointerException.class, () -> new ExcelDataValidationPrompt(null, "Text", true));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelDataValidationPrompt(" ", "Text", true));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelDataValidationPrompt("Title", " ", true));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelDataValidationErrorAlert(null, "Title", "Text", true));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDataValidationErrorAlert(
                ExcelDataValidationErrorStyle.STOP, " ", "Text", true));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDataValidationErrorAlert(
                ExcelDataValidationErrorStyle.STOP, "Title", " ", true));
  }

  @Test
  void explicitListCopiesValuesAndRejectsInvalidCollections() {
    List<String> values = new ArrayList<>(List.of("Queued", "Done"));

    ExcelDataValidationRule.ExplicitList explicitList =
        new ExcelDataValidationRule.ExplicitList(values);
    values.clear();

    assertEquals(List.of("Queued", "Done"), explicitList.values());
    assertEquals(List.of(), new ExcelDataValidationRule.ExplicitList(List.of()).values());
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelDataValidationRule.ExplicitList(List.of("Queued", " ")));
  }

  @Test
  void listAndCustomRulesNormalizeLeadingEqualsSigns() {
    assertEquals(
        new ExcelDataValidationRule.FormulaList("Statuses"),
        new ExcelDataValidationRule.FormulaList("=Statuses"));
    assertEquals(
        new ExcelDataValidationRule.CustomFormula("LEN(A1)>0"),
        new ExcelDataValidationRule.CustomFormula("=LEN(A1)>0"));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelDataValidationRule.FormulaList(" = "));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelDataValidationRule.CustomFormula(" "));
  }

  @Test
  void comparisonRulesNormalizeFormulasAcrossFamilies() {
    assertEquals(
        new ExcelDataValidationRule.WholeNumber(
            ExcelComparisonOperator.BETWEEN, "1", Optional.of("10")),
        new ExcelDataValidationRule.WholeNumber(
            ExcelComparisonOperator.BETWEEN, "=1", Optional.of(" =10 ")));
    assertEquals(
        new ExcelDataValidationRule.WholeNumber(
            ExcelComparisonOperator.NOT_BETWEEN, "1", Optional.of("10")),
        new ExcelDataValidationRule.WholeNumber(
            ExcelComparisonOperator.NOT_BETWEEN, "=1", Optional.of(" =10 ")));
    assertEquals(
        new ExcelDataValidationRule.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN, "0.5", Optional.empty()),
        new ExcelDataValidationRule.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN, "=0.5", Optional.empty()));
    assertEquals(
        new ExcelDataValidationRule.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN, "0.5", Optional.empty()),
        new ExcelDataValidationRule.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN, "=0.5", Optional.of("   ")));
    assertEquals(
        new ExcelDataValidationRule.DateRule(
            ExcelComparisonOperator.EQUAL, "DATE(2026,4,1)", Optional.empty()),
        new ExcelDataValidationRule.DateRule(
            ExcelComparisonOperator.EQUAL, "=DATE(2026,4,1)", Optional.empty()));
    assertEquals(
        new ExcelDataValidationRule.TimeRule(
            ExcelComparisonOperator.GREATER_THAN, "TIME(9,0,0)", Optional.empty()),
        new ExcelDataValidationRule.TimeRule(
            ExcelComparisonOperator.GREATER_THAN, "=TIME(9,0,0)", Optional.empty()));
    assertEquals(
        new ExcelDataValidationRule.TextLength(
            ExcelComparisonOperator.LESS_OR_EQUAL, "20", Optional.empty()),
        new ExcelDataValidationRule.TextLength(
            ExcelComparisonOperator.LESS_OR_EQUAL, "=20", Optional.empty()));
  }

  @Test
  void comparisonRulesRejectInvalidOperandShapes() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDataValidationRule.WholeNumber(
                ExcelComparisonOperator.BETWEEN, "1", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDataValidationRule.DecimalNumber(
                ExcelComparisonOperator.GREATER_THAN, "1", Optional.of("2")));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelDataValidationRule.DateRule(null, "DATE(2026,4,1)", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDataValidationRule.TimeRule(
                ExcelComparisonOperator.EQUAL, " ", Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDataValidationRule.TextLength(
                ExcelComparisonOperator.NOT_BETWEEN, "1", Optional.of(" ")));
  }

  @Test
  void snapshotVariantsValidateRangesAndFields() {
    ExcelDataValidationDefinition definition =
        new ExcelDataValidationDefinition(
            new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done")),
            true,
            false,
            Optional.of(
                new ExcelDataValidationPrompt("Status", "Choose one workflow state.", true)),
            Optional.empty());

    ExcelDataValidationSnapshot.Supported supported =
        new ExcelDataValidationSnapshot.Supported(List.of("A1:A3"), definition);
    ExcelDataValidationSnapshot.Unsupported unsupported =
        new ExcelDataValidationSnapshot.Unsupported(List.of("B1:B3"), "ANY", "Not modeled");
    ExcelDataValidationSnapshot.Unsupported invalid =
        new ExcelDataValidationSnapshot.Unsupported(
            List.of("C1:C3"), "MISSING_FORMULA", "Missing formula1.");

    assertEquals(definition, supported.validation());
    assertEquals("ANY", unsupported.kind());
    assertEquals("Missing formula1.", invalid.detail());
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelDataValidationSnapshot.Supported(List.of(), definition));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelDataValidationSnapshot.Unsupported(List.of(" "), "ANY", "detail"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelDataValidationSnapshot.Unsupported(List.of("A1"), " ", "detail"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelDataValidationSnapshot.Unsupported(List.of("A1"), "KIND", " "));
  }
}
