package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Regression coverage for formula capability assessment domain types. */
class ExcelFormulaCapabilityAssessmentTest {
  @Test
  void validatesAssessmentInvariants() {
    ExcelFormulaCapabilityAssessment evaluable =
        new ExcelFormulaCapabilityAssessment(
            "Budget", "B1", "A1*2", ExcelFormulaCapabilityKind.EVALUABLE_NOW, null, null);

    assertEquals("Budget", evaluable.sheetName());
    assertEquals("B1", evaluable.address());
    assertEquals("A1*2", evaluable.formula());
    assertEquals(ExcelFormulaCapabilityKind.EVALUABLE_NOW, evaluable.capability());
    assertNull(evaluable.issue());
    assertNull(evaluable.message());

    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelFormulaCapabilityAssessment(
                null, "B1", "A1*2", ExcelFormulaCapabilityKind.EVALUABLE_NOW, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelFormulaCapabilityAssessment(
                " ", "B1", "A1*2", ExcelFormulaCapabilityKind.EVALUABLE_NOW, null, null));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelFormulaCapabilityAssessment(
                "Budget", null, "A1*2", ExcelFormulaCapabilityKind.EVALUABLE_NOW, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelFormulaCapabilityAssessment(
                "Budget", " ", "A1*2", ExcelFormulaCapabilityKind.EVALUABLE_NOW, null, null));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelFormulaCapabilityAssessment(
                "Budget", "B1", null, ExcelFormulaCapabilityKind.EVALUABLE_NOW, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelFormulaCapabilityAssessment(
                "Budget", "B1", " ", ExcelFormulaCapabilityKind.EVALUABLE_NOW, null, null));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelFormulaCapabilityAssessment("Budget", "B1", "A1*2", null, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelFormulaCapabilityAssessment(
                "Budget",
                "B1",
                "A1*2",
                ExcelFormulaCapabilityKind.EVALUABLE_NOW,
                ExcelFormulaCapabilityIssue.INVALID_FORMULA,
                "unexpected"));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelFormulaCapabilityAssessment(
                "Budget", "B1", "A1*2", ExcelFormulaCapabilityKind.UNEVALUABLE_NOW, null, null));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelFormulaCapabilityAssessment(
                "Budget",
                "B1",
                "A1*2",
                ExcelFormulaCapabilityKind.UNPARSEABLE_BY_POI,
                ExcelFormulaCapabilityIssue.INVALID_FORMULA,
                " "));
  }

  @Test
  void exposesStableEnumNames() {
    assertEquals(
        ExcelFormulaCapabilityKind.EVALUABLE_NOW,
        ExcelFormulaCapabilityKind.valueOf("EVALUABLE_NOW"));
    assertEquals(
        ExcelFormulaCapabilityKind.UNEVALUABLE_NOW,
        ExcelFormulaCapabilityKind.valueOf("UNEVALUABLE_NOW"));
    assertEquals(
        ExcelFormulaCapabilityKind.UNPARSEABLE_BY_POI,
        ExcelFormulaCapabilityKind.valueOf("UNPARSEABLE_BY_POI"));

    assertEquals(
        ExcelFormulaCapabilityIssue.INVALID_FORMULA,
        ExcelFormulaCapabilityIssue.valueOf("INVALID_FORMULA"));
    assertEquals(
        ExcelFormulaCapabilityIssue.MISSING_EXTERNAL_WORKBOOK,
        ExcelFormulaCapabilityIssue.valueOf("MISSING_EXTERNAL_WORKBOOK"));
    assertEquals(
        ExcelFormulaCapabilityIssue.UNREGISTERED_USER_DEFINED_FUNCTION,
        ExcelFormulaCapabilityIssue.valueOf("UNREGISTERED_USER_DEFINED_FUNCTION"));
    assertEquals(
        ExcelFormulaCapabilityIssue.UNSUPPORTED_FORMULA,
        ExcelFormulaCapabilityIssue.valueOf("UNSUPPORTED_FORMULA"));
  }
}
