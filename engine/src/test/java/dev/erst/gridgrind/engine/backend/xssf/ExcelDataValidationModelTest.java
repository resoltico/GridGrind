package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.validation.DataValidationEntryReport;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.engine.conditionalformatting.ConditionalFormattingAdapter;
import dev.erst.gridgrind.engine.formulas.ComparisonFormulas;
import dev.erst.gridgrind.engine.validation.DataValidationAdapter;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import org.apache.poi.ss.usermodel.ComparisonOperator;
import org.apache.poi.ss.usermodel.DataValidation;
import org.junit.jupiter.api.Test;

/** Tests for data-validation enums and immutable model records. */
class ExcelDataValidationModelTest {
  @Test
  void poiStorageNormalizesRuleFormulasBeforeReplacingExistingValidation() throws Exception {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      var sheet = workbook.getOrCreateSheet("Rules");
      sheet
          .metadata()
          .setDataValidation(
              "A1",
              new ResolvedDataValidation(
                  new DataValidationRuleSpec.WholeNumber(
                      ExcelComparisonOperator.BETWEEN, " =1 ", Optional.of(" =10 ")),
                  false,
                  false,
                  Optional.empty(),
                  Optional.empty()));
      var validations = workbook.xssfWorkbook().getSheet("Rules").getDataValidations();
      assertEquals("1", validations.getFirst().getValidationConstraint().getFormula1());
      assertEquals("10", validations.getFirst().getValidationConstraint().getFormula2());

      assertThrows(
          IllegalArgumentException.class,
          () ->
              sheet
                  .metadata()
                  .setDataValidation(
                      "A1",
                      new ResolvedDataValidation(
                          new DataValidationRuleSpec.FormulaList(" = "),
                          false,
                          false,
                          Optional.empty(),
                          Optional.empty())));
      validations = workbook.xssfWorkbook().getSheet("Rules").getDataValidations();
      assertEquals(1, validations.size());
      assertEquals("1", validations.getFirst().getValidationConstraint().getFormula1());
      assertEquals("10", validations.getFirst().getValidationConstraint().getFormula2());
    }
  }

  @Test
  void comparisonOperatorRoundTripsPoiConstants() {
    assertEquals(
        ExcelComparisonOperator.BETWEEN,
        ConditionalFormattingAdapter.fromPoi(ComparisonOperator.BETWEEN));
    assertEquals(
        ExcelComparisonOperator.NOT_BETWEEN,
        ConditionalFormattingAdapter.fromPoi(ComparisonOperator.NOT_BETWEEN));
    assertEquals(
        ExcelComparisonOperator.EQUAL,
        ConditionalFormattingAdapter.fromPoi(ComparisonOperator.EQUAL));
    assertEquals(
        ExcelComparisonOperator.NOT_EQUAL,
        ConditionalFormattingAdapter.fromPoi(ComparisonOperator.NOT_EQUAL));
    assertEquals(
        ExcelComparisonOperator.GREATER_THAN,
        ConditionalFormattingAdapter.fromPoi(ComparisonOperator.GT));
    assertEquals(
        ExcelComparisonOperator.LESS_THAN,
        ConditionalFormattingAdapter.fromPoi(ComparisonOperator.LT));
    assertEquals(
        ExcelComparisonOperator.GREATER_OR_EQUAL,
        ConditionalFormattingAdapter.fromPoi(ComparisonOperator.GE));
    assertEquals(
        ExcelComparisonOperator.LESS_OR_EQUAL,
        ConditionalFormattingAdapter.fromPoi(ComparisonOperator.LE));
    assertEquals(
        ComparisonOperator.BETWEEN,
        ConditionalFormattingAdapter.toPoi(ExcelComparisonOperator.BETWEEN));
    assertEquals(
        ComparisonOperator.NOT_BETWEEN,
        ConditionalFormattingAdapter.toPoi(ExcelComparisonOperator.NOT_BETWEEN));
    assertEquals(
        ComparisonOperator.EQUAL,
        ConditionalFormattingAdapter.toPoi(ExcelComparisonOperator.EQUAL));
    assertEquals(
        ComparisonOperator.NOT_EQUAL,
        ConditionalFormattingAdapter.toPoi(ExcelComparisonOperator.NOT_EQUAL));
    assertEquals(
        ComparisonOperator.GT,
        ConditionalFormattingAdapter.toPoi(ExcelComparisonOperator.GREATER_THAN));
    assertEquals(
        ComparisonOperator.LT,
        ConditionalFormattingAdapter.toPoi(ExcelComparisonOperator.LESS_THAN));
    assertEquals(
        ComparisonOperator.GE,
        ConditionalFormattingAdapter.toPoi(ExcelComparisonOperator.GREATER_OR_EQUAL));
    assertEquals(
        ComparisonOperator.LE,
        ConditionalFormattingAdapter.toPoi(ExcelComparisonOperator.LESS_OR_EQUAL));
  }

  @Test
  void comparisonOperatorRejectsUnsupportedPoiValues() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class, () -> ConditionalFormattingAdapter.fromPoi(-99));

    assertEquals(
        "Unsupported Apache POI conditional-formatting comparison operator: -99",
        failure.getMessage());
  }

  @Test
  void errorStyleRoundTripsPoiConstants() {
    assertEquals(
        ExcelDataValidationErrorStyle.STOP,
        DataValidationAdapter.fromPoiErrorStyle(DataValidation.ErrorStyle.STOP));
    assertEquals(
        ExcelDataValidationErrorStyle.WARNING,
        DataValidationAdapter.fromPoiErrorStyle(DataValidation.ErrorStyle.WARNING));
    assertEquals(
        ExcelDataValidationErrorStyle.INFORMATION,
        DataValidationAdapter.fromPoiErrorStyle(DataValidation.ErrorStyle.INFO));
    assertEquals(
        DataValidation.ErrorStyle.STOP,
        DataValidationAdapter.toPoiErrorStyle(ExcelDataValidationErrorStyle.STOP));
    assertEquals(
        DataValidation.ErrorStyle.WARNING,
        DataValidationAdapter.toPoiErrorStyle(ExcelDataValidationErrorStyle.WARNING));
    assertEquals(
        DataValidation.ErrorStyle.INFO,
        DataValidationAdapter.toPoiErrorStyle(ExcelDataValidationErrorStyle.INFORMATION));
  }

  @Test
  void errorStyleRejectsUnsupportedPoiValues() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class, () -> DataValidationAdapter.fromPoiErrorStyle(-77));

    assertEquals("Unsupported Apache POI error style: -77", failure.getMessage());
  }

  @Test
  void promptAndAlertRejectBlankValues() {
    assertThrows(
        NullPointerException.class, () -> new ResolvedValidationPrompt(null, "Text", true));
    assertThrows(
        IllegalArgumentException.class, () -> new ResolvedValidationPrompt(" ", "Text", true));
    assertThrows(
        IllegalArgumentException.class, () -> new ResolvedValidationPrompt("Title", " ", true));
    assertThrows(
        NullPointerException.class,
        () -> new ResolvedValidationErrorAlert(null, "Title", "Text", true));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedValidationErrorAlert(
                ExcelDataValidationErrorStyle.STOP, " ", "Text", true));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ResolvedValidationErrorAlert(
                ExcelDataValidationErrorStyle.STOP, "Title", " ", true));
  }

  @Test
  void explicitListCopiesValuesAndRejectsInvalidCollections() {
    List<String> values = new ArrayList<>(List.of("Queued", "Done"));

    DataValidationRuleSpec.ExplicitList explicitList =
        new DataValidationRuleSpec.ExplicitList(values);
    values.clear();

    assertEquals(List.of("Queued", "Done"), explicitList.values());
    assertEquals(List.of(), new DataValidationRuleSpec.ExplicitList(List.of()).values());
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleSpec.ExplicitList(List.of("Queued", " ")));
  }

  @Test
  void listAndCustomRulesNormalizeLeadingEqualsSigns() {
    assertEquals(
        new DataValidationRuleSpec.FormulaList(
            ComparisonFormulas.normalizeFormula("Statuses", "formula")),
        new DataValidationRuleSpec.FormulaList(
            ComparisonFormulas.normalizeFormula("=Statuses", "formula")));
    assertEquals(
        new DataValidationRuleSpec.CustomFormula(
            ComparisonFormulas.normalizeFormula("LEN(A1)>0", "formula")),
        new DataValidationRuleSpec.CustomFormula(
            ComparisonFormulas.normalizeFormula("=LEN(A1)>0", "formula")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleSpec.FormulaList(
                ComparisonFormulas.normalizeFormula(" = ", "formula")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleSpec.CustomFormula(
                ComparisonFormulas.normalizeFormula(" ", "formula")));
  }

  @Test
  void comparisonRulesNormalizeFormulasAcrossFamilies() {
    assertEquals(
        new DataValidationRuleSpec.WholeNumber(
            ExcelComparisonOperator.BETWEEN,
            ComparisonFormulas.normalizeFormula("1", "formula1"),
            ComparisonFormulas.normalizeOptionalFormula(Optional.of("10"), "formula2")),
        new DataValidationRuleSpec.WholeNumber(
            ExcelComparisonOperator.BETWEEN,
            ComparisonFormulas.normalizeFormula("=1", "formula1"),
            ComparisonFormulas.normalizeOptionalFormula(Optional.of(" =10 "), "formula2")));
    assertEquals(
        new DataValidationRuleSpec.WholeNumber(
            ExcelComparisonOperator.NOT_BETWEEN,
            ComparisonFormulas.normalizeFormula("1", "formula1"),
            ComparisonFormulas.normalizeOptionalFormula(Optional.of("10"), "formula2")),
        new DataValidationRuleSpec.WholeNumber(
            ExcelComparisonOperator.NOT_BETWEEN,
            ComparisonFormulas.normalizeFormula("=1", "formula1"),
            ComparisonFormulas.normalizeOptionalFormula(Optional.of(" =10 "), "formula2")));
    assertEquals(
        new DataValidationRuleSpec.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN,
            ComparisonFormulas.normalizeFormula("0.5", "formula1"),
            ComparisonFormulas.normalizeOptionalFormula(Optional.empty(), "formula2")),
        new DataValidationRuleSpec.DecimalNumber(
            ExcelComparisonOperator.GREATER_THAN,
            ComparisonFormulas.normalizeFormula("=0.5", "formula1"),
            ComparisonFormulas.normalizeOptionalFormula(Optional.empty(), "formula2")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleSpec.DecimalNumber(
                ExcelComparisonOperator.GREATER_THAN, "0.5", Optional.of("   ")));
    assertEquals(
        new DataValidationRuleSpec.DateRule(
            ExcelComparisonOperator.EQUAL,
            ComparisonFormulas.normalizeFormula("DATE(2026,4,1)", "formula1"),
            ComparisonFormulas.normalizeOptionalFormula(Optional.empty(), "formula2")),
        new DataValidationRuleSpec.DateRule(
            ExcelComparisonOperator.EQUAL,
            ComparisonFormulas.normalizeFormula("=DATE(2026,4,1)", "formula1"),
            ComparisonFormulas.normalizeOptionalFormula(Optional.empty(), "formula2")));
    assertEquals(
        new DataValidationRuleSpec.TimeRule(
            ExcelComparisonOperator.GREATER_THAN,
            ComparisonFormulas.normalizeFormula("TIME(9,0,0)", "formula1"),
            ComparisonFormulas.normalizeOptionalFormula(Optional.empty(), "formula2")),
        new DataValidationRuleSpec.TimeRule(
            ExcelComparisonOperator.GREATER_THAN,
            ComparisonFormulas.normalizeFormula("=TIME(9,0,0)", "formula1"),
            ComparisonFormulas.normalizeOptionalFormula(Optional.empty(), "formula2")));
    assertEquals(
        new DataValidationRuleSpec.TextLength(
            ExcelComparisonOperator.LESS_OR_EQUAL,
            ComparisonFormulas.normalizeFormula("20", "formula1"),
            ComparisonFormulas.normalizeOptionalFormula(Optional.empty(), "formula2")),
        new DataValidationRuleSpec.TextLength(
            ExcelComparisonOperator.LESS_OR_EQUAL,
            ComparisonFormulas.normalizeFormula("=20", "formula1"),
            ComparisonFormulas.normalizeOptionalFormula(Optional.empty(), "formula2")));
  }

  @Test
  void comparisonRulesRejectInvalidOperandShapes() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleSpec.WholeNumber(
                ExcelComparisonOperator.BETWEEN,
                ComparisonFormulas.normalizeFormula("1", "formula1"),
                ComparisonFormulas.normalizeOptionalFormula(Optional.empty(), "formula2")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleSpec.DecimalNumber(
                ExcelComparisonOperator.GREATER_THAN,
                ComparisonFormulas.normalizeFormula("1", "formula1"),
                ComparisonFormulas.normalizeOptionalFormula(Optional.of("2"), "formula2")));
    assertThrows(
        NullPointerException.class,
        () ->
            new DataValidationRuleSpec.DateRule(
                null,
                ComparisonFormulas.normalizeFormula("DATE(2026,4,1)", "formula1"),
                ComparisonFormulas.normalizeOptionalFormula(Optional.empty(), "formula2")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleSpec.TimeRule(
                ExcelComparisonOperator.EQUAL,
                ComparisonFormulas.normalizeFormula(" ", "formula1"),
                ComparisonFormulas.normalizeOptionalFormula(Optional.empty(), "formula2")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleSpec.TextLength(
                ExcelComparisonOperator.NOT_BETWEEN,
                ComparisonFormulas.normalizeFormula("1", "formula1"),
                ComparisonFormulas.normalizeOptionalFormula(Optional.of(" "), "formula2")));
  }

  @Test
  void snapshotVariantsValidateRangesAndFields() {
    ResolvedDataValidation definition =
        new ResolvedDataValidation(
            new DataValidationRuleSpec.ExplicitList(List.of("Queued", "Done")),
            true,
            false,
            Optional.of(new ResolvedValidationPrompt("Status", "Choose one workflow state.", true)),
            Optional.empty());

    DataValidationEntryReport.Supported supported =
        new DataValidationEntryReport.Supported(
            List.of("A1:A3").stream()
                .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                .toList(),
            dev.erst.gridgrind.engine.backend.xssf.DataValidationFixtures.report(definition));
    DataValidationEntryReport.Unsupported unsupported =
        new DataValidationEntryReport.Unsupported(
            List.of("B1:B3").stream()
                .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                .toList(),
            "ANY",
            "Not modeled");
    DataValidationEntryReport.Unsupported invalid =
        new DataValidationEntryReport.Unsupported(
            List.of("C1:C3").stream()
                .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                .toList(),
            "MISSING_FORMULA",
            "Missing formula1.");

    assertEquals(
        dev.erst.gridgrind.engine.backend.xssf.DataValidationFixtures.report(definition),
        supported.validation());
    assertEquals("ANY", unsupported.kind());
    assertEquals("Missing formula1.", invalid.detail());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationEntryReport.Supported(
                List.<String>of().stream()
                    .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                    .toList(),
                dev.erst.gridgrind.engine.backend.xssf.DataValidationFixtures.report(definition)));
    assertDoesNotThrow(
        () ->
            new DataValidationEntryReport.Unsupported(
                List.of(" ").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                    .toList(),
                "ANY",
                "detail"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationEntryReport.Unsupported(
                List.of("A1").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                    .toList(),
                " ",
                "detail"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationEntryReport.Unsupported(
                List.of("A1").stream()
                    .map(dev.erst.gridgrind.core.vocabulary.RangeLabel::new)
                    .toList(),
                "KIND",
                " "));
  }
}
