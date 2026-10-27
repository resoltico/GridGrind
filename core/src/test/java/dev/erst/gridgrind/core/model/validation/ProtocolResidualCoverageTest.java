package dev.erst.gridgrind.core.model.validation;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.chart.ChartTitleSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.ShapeSpec;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaSpec;
import dev.erst.gridgrind.core.model.formula.FormulaUdfFunctionSpec;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.layout.WindowRowReport;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.DifferentialStylePatch;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Closes focused DTO factory and validation coverage gaps left by broader protocol tests. */
class ProtocolResidualCoverageTest {
  @Test
  void colorFactoriesAndOptionalValidationBranchesStayExplicit() {
    assertEquals(Optional.of(0.25d), CellColorReport.rgb(new RgbHex("#112233"), 0.25d).tint());
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
                    new DifferentialStylePatch(
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
    ConditionalFormattingRuleSpec.CellValueRule betweenRule =
        new ConditionalFormattingRuleSpec.CellValueRule(
            ExcelComparisonOperator.BETWEEN, "1", Optional.of(" 9 "), true, Optional.empty());
    ConditionalFormattingRuleSpec.CellValueRule notBetweenRule =
        new ConditionalFormattingRuleSpec.CellValueRule(
            ExcelComparisonOperator.NOT_BETWEEN, "1", Optional.of(" 11 "), true, Optional.empty());

    assertEquals(Optional.of("9"), betweenRule.formula2());
    assertEquals(Optional.of("11"), notBetweenRule.formula2());
    assertEquals(
        "rank must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new ConditionalFormattingRuleSpec.Top10Rule(
                        false, 0, false, false, Optional.empty()))
            .getMessage());
    assertEquals(
        "formula2 must be omitted unless operator is BETWEEN or NOT_BETWEEN",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DataValidationRuleSpec.WholeNumber(
                        ExcelComparisonOperator.GREATER_THAN, "7", Optional.of("9")))
            .getMessage());
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleSpec.CustomFormula("DDE(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleSpec.FormulaList("=dde(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleSpec.WholeNumber(
                ExcelComparisonOperator.BETWEEN,
                "1",
                Optional.of("DDE(\"cmd\",\"/C calc\",\"\")")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleSpec.WholeNumber(
                ExcelComparisonOperator.NOT_BETWEEN,
                "1",
                Optional.of("DDE(\"cmd\",\"/C calc\",\"\")")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleSpec.FormulaRule(
                "DDE(\"cmd\",\"/C calc\",\"\")", false, Optional.empty()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleSpec.CellValueRule(
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
        new DataValidationRuleSpec.WholeNumber(
                ExcelComparisonOperator.BETWEEN, "1", Optional.of("9"))
            .formula2());
    assertEquals(
        Optional.of("9"),
        new DataValidationRuleSpec.DecimalNumber(
                ExcelComparisonOperator.BETWEEN, "1", Optional.of("9"))
            .formula2());
    assertEquals(
        Optional.of("9"),
        new DataValidationRuleSpec.DateRule(ExcelComparisonOperator.BETWEEN, "1", Optional.of("9"))
            .formula2());
    assertEquals(
        Optional.of("9"),
        new DataValidationRuleSpec.TimeRule(ExcelComparisonOperator.BETWEEN, "1", Optional.of("9"))
            .formula2());
    assertEquals(
        Optional.of("9"),
        new DataValidationRuleSpec.TextLength(
                ExcelComparisonOperator.BETWEEN, "1", Optional.of("9"))
            .formula2());
    assertEquals(
        Optional.empty(),
        new DataValidationRuleSpec.WholeNumber(
                ExcelComparisonOperator.GREATER_THAN, "1", Optional.empty())
            .formula2());
    assertEquals(
        Optional.of("9"),
        new DataValidationRuleSpec.WholeNumber(
                ExcelComparisonOperator.NOT_BETWEEN, "1", Optional.of("9"))
            .formula2());
    assertEquals("safe", new DataValidationRuleSpec.FormulaList("safe").formula());
    assertEquals("safe", new DataValidationRuleSpec.CustomFormula("safe").formula());
    assertEquals(
        "formula2 must not be blank for between",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new DataValidationRuleSpec.WholeNumber(
                        ExcelComparisonOperator.BETWEEN, "1", Optional.empty()))
            .getMessage());
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleSpec.WholeNumber(
                ExcelComparisonOperator.BETWEEN,
                "DDE(\"cmd\",\"/C calc\",\"\")",
                Optional.of("9")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleSpec.DecimalNumber(
                ExcelComparisonOperator.BETWEEN,
                "DDE(\"cmd\",\"/C calc\",\"\")",
                Optional.of("9")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleSpec.DateRule(
                ExcelComparisonOperator.BETWEEN,
                "DDE(\"cmd\",\"/C calc\",\"\")",
                Optional.of("9")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleSpec.TimeRule(
                ExcelComparisonOperator.BETWEEN,
                "DDE(\"cmd\",\"/C calc\",\"\")",
                Optional.of("9")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new DataValidationRuleSpec.TextLength(
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
                () -> new DataValidationRuleSpec.ExplicitList(List.of("x".repeat(256), "y")))
            .getMessage());
    assertDoesNotThrow(() -> new DataValidationRuleSpec.ExplicitList(List.of("x".repeat(253))));
    assertEquals(
        "values must serialize to at most 255 characters for Excel data validation; got 256",
        assertThrows(
                IllegalArgumentException.class,
                () -> new DataValidationRuleSpec.ExplicitList(List.of("x".repeat(254))))
            .getMessage());
    assertEquals(
        "values must not contain ',' or '\"' in an explicit list; use FORMULA_LIST for values that require delimiters",
        assertThrows(
                IllegalArgumentException.class,
                () -> new DataValidationRuleSpec.ExplicitList(List.of("North, East")))
            .getMessage());
    assertEquals(
        "values must not contain ',' or '\"' in an explicit list; use FORMULA_LIST for values that require delimiters",
        assertThrows(
                IllegalArgumentException.class,
                () -> new DataValidationRuleSpec.ExplicitList(List.of("North\" East")))
            .getMessage());
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleSpec.ExplicitList(List.of(",North")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleSpec.ExplicitList(List.of("\"North")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleSpec.ExplicitList(List.of(" ")));
    assertEquals(
        List.of("North", "West"),
        new DataValidationRuleSpec.ExplicitList(List.of("North", "West")).values());
  }

  @Test
  void rejectsWebserviceFormulaInjectionInAllCaseForms() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellValueSpec.Formula(
                TextSourceSpec.inline("WEBSERVICE(\"https://evil.example.com\")")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new CellValueSpec.Formula(
                TextSourceSpec.inline("=webservice(\"https://evil.example.com\")")));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ArrayFormulaSpec(
                TextSourceSpec.inline("{=WEBSERVICE(\"https://evil.example.com\")}")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new DataValidationRuleSpec.CustomFormula("WEBSERVICE(\"https://evil.example.com\")"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ConditionalFormattingRuleSpec.FormulaRule(
                "WEBSERVICE(\"https://evil.example.com\")", false, Optional.empty()));
  }

  @Test
  void rejectsFormulaInjectionInNamedRangeChartTitleAndUdfTemplate() {
    // Named-range formula target
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

    // Chart title formula
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartTitleSpec.Formula("DDE(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ChartTitleSpec.Formula("WEBSERVICE(\"https://evil.example.com\")"));

    // UDF formula template (defense-in-depth; POI evaluator won't execute these
    //           server-side, but they must not silently end up in future execution paths)
    assertThrows(
        IllegalArgumentException.class,
        () -> new FormulaUdfFunctionSpec("MY_FUNC", 0, 0, "DDE(\"cmd\",\"/C calc\",\"\")"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new FormulaUdfFunctionSpec(
                "MY_FUNC", 0, 0, "WEBSERVICE(\"https://evil.example.com\")"));
  }

  @Test
  void simpleShapeKindRemainsExplicit() {
    DrawingAnchorSpec.TwoCell anchor =
        new DrawingAnchorSpec.TwoCell(
            new DrawingMarkerSpec(0, 0),
            new DrawingMarkerSpec(2, 2),
            ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
    ShapeSpec.SimpleShape shape =
        new ShapeSpec.SimpleShape(
            "Queue Banner", anchor, "rect", Optional.of(TextSourceSpec.inline("Ready")));

    assertEquals(ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE, shape.kind());
  }
}
