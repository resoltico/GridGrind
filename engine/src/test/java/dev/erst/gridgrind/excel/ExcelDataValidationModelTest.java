package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.DataValidation;
import org.junit.jupiter.api.Test;

/** Tests for data-validation enums and immutable model records. */
class ExcelDataValidationModelTest {
  @Test
  void comparisonOperatorRoundTripsPoiConstants() {
    assertEquals(
        ExcelComparisonOperator.BETWEEN,
        ExcelComparisonOperatorPoiBridge.fromPoi(ComparisonOperator.BETWEEN));
    assertEquals(
        ExcelComparisonOperator.NOT_BETWEEN,
        ExcelComparisonOperatorPoiBridge.fromPoi(ComparisonOperator.NOT_BETWEEN));
    assertEquals(
        ExcelComparisonOperator.EQUAL,
        ExcelComparisonOperatorPoiBridge.fromPoi(ComparisonOperator.EQUAL));
    assertEquals(
        ExcelComparisonOperator.NOT_EQUAL,
        ExcelComparisonOperatorPoiBridge.fromPoi(ComparisonOperator.NOT_EQUAL));
    assertEquals(
        ExcelComparisonOperator.GREATER_THAN,
        ExcelComparisonOperatorPoiBridge.fromPoi(ComparisonOperator.GT));
    assertEquals(
        ExcelComparisonOperator.LESS_THAN,
        ExcelComparisonOperatorPoiBridge.fromPoi(ComparisonOperator.LT));
    assertEquals(
        ExcelComparisonOperator.GREATER_OR_EQUAL,
        ExcelComparisonOperatorPoiBridge.fromPoi(ComparisonOperator.GE));
    assertEquals(
        ExcelComparisonOperator.LESS_OR_EQUAL,
        ExcelComparisonOperatorPoiBridge.fromPoi(ComparisonOperator.LE));
    assertEquals(
        ComparisonOperator.BETWEEN,
        ExcelComparisonOperatorPoiBridge.toPoi(ExcelComparisonOperator.BETWEEN));
    assertEquals(
        ComparisonOperator.NOT_BETWEEN,
        ExcelComparisonOperatorPoiBridge.toPoi(ExcelComparisonOperator.NOT_BETWEEN));
    assertEquals(
        ComparisonOperator.EQUAL,
        ExcelComparisonOperatorPoiBridge.toPoi(ExcelComparisonOperator.EQUAL));
    assertEquals(
        ComparisonOperator.NOT_EQUAL,
        ExcelComparisonOperatorPoiBridge.toPoi(ExcelComparisonOperator.NOT_EQUAL));
    assertEquals(
        ComparisonOperator.GT,
        ExcelComparisonOperatorPoiBridge.toPoi(ExcelComparisonOperator.GREATER_THAN));
    assertEquals(
        ComparisonOperator.LT,
        ExcelComparisonOperatorPoiBridge.toPoi(ExcelComparisonOperator.LESS_THAN));
    assertEquals(
        ComparisonOperator.GE,
        ExcelComparisonOperatorPoiBridge.toPoi(ExcelComparisonOperator.GREATER_OR_EQUAL));
    assertEquals(
        ComparisonOperator.LE,
        ExcelComparisonOperatorPoiBridge.toPoi(ExcelComparisonOperator.LESS_OR_EQUAL));
  }

  @Test
  void comparisonOperatorRejectsUnsupportedPoiValues() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class, () -> ExcelComparisonOperatorPoiBridge.fromPoi(-99));

    assertEquals("Unsupported Apache POI comparison operator: -99", failure.getMessage());
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
        new ExcelDataValidationRule.WholeNumber(ExcelComparisonOperator.BETWEEN, "1", "10"),
        new ExcelDataValidationRule.WholeNumber(ExcelComparisonOperator.BETWEEN, "=1", " =10 "));
    assertEquals(
        new ExcelDataValidationRule.WholeNumber(ExcelComparisonOperator.NOT_BETWEEN, "1", "10"),
        new ExcelDataValidationRule.WholeNumber(
            ExcelComparisonOperator.NOT_BETWEEN, "=1", " =10 "));
    assertEquals(
        new ExcelDataValidationRule.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN, "0.5", null),
        new ExcelDataValidationRule.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN, "=0.5", null));
    assertEquals(
        new ExcelDataValidationRule.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN, "0.5", null),
        new ExcelDataValidationRule.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN, "=0.5", "   "));
    assertEquals(
        new ExcelDataValidationRule.DateRule(ExcelComparisonOperator.EQUAL, "DATE(2026,4,1)", null),
        new ExcelDataValidationRule.DateRule(
            ExcelComparisonOperator.EQUAL, "=DATE(2026,4,1)", null));
    assertEquals(
        new ExcelDataValidationRule.TimeRule(
            ExcelComparisonOperator.GREATER_THAN, "TIME(9,0,0)", null),
        new ExcelDataValidationRule.TimeRule(
            ExcelComparisonOperator.GREATER_THAN, "=TIME(9,0,0)", null));
    assertEquals(
        new ExcelDataValidationRule.TextLength(ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
        new ExcelDataValidationRule.TextLength(ExcelComparisonOperator.LESS_OR_EQUAL, "=20", null));
  }

  @Test
  void comparisonRulesRejectInvalidOperandShapes() {
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelDataValidationRule.WholeNumber(ExcelComparisonOperator.BETWEEN, "1", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDataValidationRule.DecimalNumber(
                ExcelComparisonOperator.GREATER_THAN, "1", "2"));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelDataValidationRule.DateRule(null, "DATE(2026,4,1)", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelDataValidationRule.TimeRule(ExcelComparisonOperator.EQUAL, " ", null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelDataValidationRule.TextLength(ExcelComparisonOperator.NOT_BETWEEN, "1", " "));
  }

  @Test
  void snapshotVariantsValidateRangesAndFields() {
    ExcelDataValidationDefinition definition =
        new ExcelDataValidationDefinition(
            new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done")),
            true,
            false,
            new ExcelDataValidationPrompt("Status", "Choose one workflow state.", true),
            null);

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
