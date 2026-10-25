package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Closes focused DTO factory and validation coverage gaps left by broader protocol tests. */
class ProtocolResidualCoverageTest {
  @Test
  void colorFactoriesAndOptionalValidationBranchesStayExplicit() {
    assertEquals(Optional.of(0.25d), CellColorReport.rgb("#112233", 0.25d).tint());
    assertEquals(Optional.of(-0.5d), CellColorReport.indexed(7, -0.5d).tint());
    assertEquals(
        "rowIndex must not be negative",
        assertThrows(IllegalArgumentException.class, () -> new WindowRowReport(-1, List.of()))
            .getMessage());
    assertEquals(
        "style must set at least one attribute",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DifferentialStyleInput(
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty(),
                        Optional.empty()))
            .getMessage());
  }

  @Test
  void conditionalFormattingAndValidationOptionalsNormalizeAtTheTypeBoundary() {
    ConditionalFormattingRuleInput.CellValueRule betweenRule =
        new ConditionalFormattingRuleInput.CellValueRule(
            ExcelComparisonOperator.BETWEEN, "1", Optional.of(" 9 "), true, Optional.empty());
    ConditionalFormattingRuleInput.CellValueRule notBetweenRule =
        new ConditionalFormattingRuleInput.CellValueRule(
            ExcelComparisonOperator.NOT_BETWEEN, "1", Optional.of(" 11 "), true, Optional.empty());

    assertEquals(Optional.of("9"), betweenRule.formula2());
    assertEquals(Optional.of("11"), notBetweenRule.formula2());
    assertEquals(
        "rank must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ConditionalFormattingRuleInput.Top10Rule(
                        false, 0, false, false, Optional.empty()))
            .getMessage());
    assertEquals(
        "formula2 must be omitted unless operator is BETWEEN or NOT_BETWEEN",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DataValidationRuleInput.WholeNumber(
                        ExcelComparisonOperator.GREATER_THAN, "7", Optional.of("9")))
            .getMessage());
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleInput.CustomFormula("DDE(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleInput.FormulaList("=dde(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleInput.WholeNumber(
                ExcelComparisonOperator.BETWEEN,
                "1",
                Optional.of("DDE(\"cmd\",\"/C calc\",\"\")")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleInput.WholeNumber(
                ExcelComparisonOperator.NOT_BETWEEN,
                "1",
                Optional.of("DDE(\"cmd\",\"/C calc\",\"\")")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.FormulaRule(
                "DDE(\"cmd\",\"/C calc\",\"\")", false, Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.CellValueRule(
                ExcelComparisonOperator.BETWEEN,
                "1",
                Optional.of("DDE(\"cmd\",\"/C calc\",\"\")"),
                false,
                Optional.empty()));
  }

  @Test
  void everyComparisonValidationSubtypeEnforcesItsCompleteFormulaContract() {
    assertEquals(
        Optional.of("9"),
        new DataValidationRuleInput.WholeNumber(
                ExcelComparisonOperator.BETWEEN, "1", Optional.of("9"))
            .formula2());
    assertEquals(
        Optional.of("9"),
        new DataValidationRuleInput.DecimalNumber(
                ExcelComparisonOperator.BETWEEN, "1", Optional.of("9"))
            .formula2());
    assertEquals(
        Optional.of("9"),
        new DataValidationRuleInput.DateRule(ExcelComparisonOperator.BETWEEN, "1", Optional.of("9"))
            .formula2());
    assertEquals(
        Optional.of("9"),
        new DataValidationRuleInput.TimeRule(ExcelComparisonOperator.BETWEEN, "1", Optional.of("9"))
            .formula2());
    assertEquals(
        Optional.of("9"),
        new DataValidationRuleInput.TextLength(
                ExcelComparisonOperator.BETWEEN, "1", Optional.of("9"))
            .formula2());
    assertEquals(
        Optional.empty(),
        new DataValidationRuleInput.WholeNumber(
                ExcelComparisonOperator.GREATER_THAN, "1", Optional.empty())
            .formula2());
    assertEquals(
        Optional.of("9"),
        new DataValidationRuleInput.WholeNumber(
                ExcelComparisonOperator.NOT_BETWEEN, "1", Optional.of("9"))
            .formula2());
    assertEquals("safe", new DataValidationRuleInput.FormulaList("safe").formula());
    assertEquals("safe", new DataValidationRuleInput.CustomFormula("safe").formula());
    assertEquals(
        "formula2 must not be blank for between",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DataValidationRuleInput.WholeNumber(
                        ExcelComparisonOperator.BETWEEN, "1", Optional.empty()))
            .getMessage());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleInput.WholeNumber(
                ExcelComparisonOperator.BETWEEN,
                "DDE(\"cmd\",\"/C calc\",\"\")",
                Optional.of("9")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleInput.DecimalNumber(
                ExcelComparisonOperator.BETWEEN,
                "DDE(\"cmd\",\"/C calc\",\"\")",
                Optional.of("9")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleInput.DateRule(
                ExcelComparisonOperator.BETWEEN,
                "DDE(\"cmd\",\"/C calc\",\"\")",
                Optional.of("9")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleInput.TimeRule(
                ExcelComparisonOperator.BETWEEN,
                "DDE(\"cmd\",\"/C calc\",\"\")",
                Optional.of("9")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleInput.TextLength(
                ExcelComparisonOperator.BETWEEN,
                "DDE(\"cmd\",\"/C calc\",\"\")",
                Optional.of("9")));
  }

  @Test
  void rejectsExplicitValidationListsThatExceedExcelsSerializedFormulaLimit() {
    assertEquals(
        "values must serialize to at most 255 characters for Excel data validation; got 260",
        assertThrows(
                IllegalArgumentException.class,
                () -> new DataValidationRuleInput.ExplicitList(List.of("x".repeat(256), "y")))
            .getMessage());
    assertDoesNotThrow(() -> new DataValidationRuleInput.ExplicitList(List.of("x".repeat(253))));
    assertEquals(
        "values must serialize to at most 255 characters for Excel data validation; got 256",
        assertThrows(
                IllegalArgumentException.class,
                () -> new DataValidationRuleInput.ExplicitList(List.of("x".repeat(254))))
            .getMessage());
    assertEquals(
        "values must not contain ',' or '\"' in an explicit list; use FORMULA_LIST for values that require delimiters",
        assertThrows(
                IllegalArgumentException.class,
                () -> new DataValidationRuleInput.ExplicitList(List.of("North, East")))
            .getMessage());
    assertEquals(
        "values must not contain ',' or '\"' in an explicit list; use FORMULA_LIST for values that require delimiters",
        assertThrows(
                IllegalArgumentException.class,
                () -> new DataValidationRuleInput.ExplicitList(List.of("North\" East")))
            .getMessage());
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleInput.ExplicitList(List.of(",North")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleInput.ExplicitList(List.of("\"North")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleInput.ExplicitList(List.of(" ")));
    assertEquals(
        List.of("North", "West"),
        new DataValidationRuleInput.ExplicitList(List.of("North", "West")).values());
  }

  @Test
  void rejectsWebserviceFormulaInjectionInAllCaseForms() { // LIM-031
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellInput.Formula(
                TextSourceInput.inline("WEBSERVICE(\"https://evil.example.com\")")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellInput.Formula(
                TextSourceInput.inline("=webservice(\"https://evil.example.com\")")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ArrayFormulaInput(
                TextSourceInput.inline("{=WEBSERVICE(\"https://evil.example.com\")}")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleInput.CustomFormula("WEBSERVICE(\"https://evil.example.com\")"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleInput.FormulaRule(
                "WEBSERVICE(\"https://evil.example.com\")", false, Optional.empty()));
  }

  @Test
  void rejectsFormulaInjectionInNamedRangeChartTitleAndUdfTemplate() { // LIM-032, LIM-033, LIM-034
    // LIM-032 — named-range formula target
    assertThrows(
        IllegalArgumentException.class,
        () -> NamedRangeTarget.formula("DDE(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class,
        () -> NamedRangeTarget.formula("=DDE(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class,
        () -> NamedRangeTarget.formula("WEBSERVICE(\"https://evil.example.com\")"));
    assertThrows(
        IllegalArgumentException.class,
        () -> NamedRangeTarget.formula("webservice(\"https://evil.example.com\")"));

    // LIM-033 — chart title formula
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartTitleInput.Formula("DDE(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartTitleInput.Formula("WEBSERVICE(\"https://evil.example.com\")"));

    // LIM-034 — UDF formula template (defense-in-depth; POI evaluator won't execute these
    //           server-side, but they must not silently end up in future execution paths)
    assertThrows(
        IllegalArgumentException.class,
        () -> new FormulaUdfFunctionInput("MY_FUNC", 0, 0, "DDE(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new FormulaUdfFunctionInput(
                "MY_FUNC", 0, 0, "WEBSERVICE(\"https://evil.example.com\")"));
  }

  @Test
  void simpleShapeKindRemainsExplicit() {
    DrawingAnchorInput.TwoCell anchor =
        new DrawingAnchorInput.TwoCell(
            new DrawingMarkerInput(0, 0),
            new DrawingMarkerInput(2, 2),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    ShapeInput.SimpleShape shape =
        new ShapeInput.SimpleShape(
            "Queue Banner", anchor, "rect", Optional.of(TextSourceInput.inline("Ready")));

    assertEquals(ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, shape.kind());
  }
}
