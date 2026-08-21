package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.json.FormulaRequestException;
import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.MutationStep;
import org.junit.jupiter.api.Test;

/** Locks the shared OOXML formula-body framing for parsed and opaque cell formulas. */
class CellFormulaFramingTest {
  @Test
  void rejectsLeadingEqualsAndEmptyFormulaBodiesForBothInputVariants() {
    assertThrows(
        InvalidFormulaInputException.class,
        () -> new CellInput.Formula(TextSourceInput.inline("=SUM(A1:A2)")));
    assertThrows(
        InvalidRawFormulaTextException.class,
        () -> new CellInput.RawFormula(TextSourceInput.inline("=LAMBDA(x,x)(A1)")));
    assertThrows(
        InvalidFormulaInputException.class,
        () -> new CellInput.Formula(TextSourceInput.inline(" ")));
    assertThrows(
        InvalidRawFormulaTextException.class,
        () -> new CellInput.RawFormula(TextSourceInput.inline(" ")));
  }

  @Test
  void enforcesDistinctNormalAndOpaqueFormulaTextContracts() {
    assertThrows(
        InvalidFormulaInputException.class,
        () -> new CellInput.Formula(TextSourceInput.inline("{=A1:A2}")));
    assertEquals(
        "{=A1",
        ((TextSourceInput.Inline) new CellInput.Formula(TextSourceInput.inline("{=A1")).source())
            .text());
    assertThrows(
        InvalidRawFormulaTextException.class,
        () -> new CellInput.RawFormula(TextSourceInput.inline("A1\u0001+B1")));
    assertThrows(
        InvalidRawFormulaTextException.class,
        () -> new CellInput.RawFormula(TextSourceInput.inline("A1\uD800+B1")));

    CellInput.RawFormula raw =
        new CellInput.RawFormula(TextSourceInput.inline("\"<b>\"&A1\t\n\r\uD83D\uDE00"));
    assertEquals("\"<b>\"&A1\t\n\r\uD83D\uDE00", ((TextSourceInput.Inline) raw.source()).text());
    FormulaTextValidation.requireXml10CharacterData(
        new String(
            new int[] {0x9, 0xA, 0xD, 0x20, 0xD7FF, 0xE000, 0xFFFD, 0x10000, 0x10FFFF}, 0, 9));
    assertThrows(
        InvalidRawFormulaTextException.class,
        () -> FormulaTextValidation.requireXml10CharacterData("A1\uFFFE+B1"));
  }

  @Test
  void preservesOpaqueFormulaBodiesWithoutApplyingNormalFormulaSecurityRules() {
    CellInput.RawFormula raw =
        new CellInput.RawFormula(TextSourceInput.inline("LAMBDA(x,x+1)(A1)"));

    assertEquals("LAMBDA(x,x+1)(A1)", ((TextSourceInput.Inline) raw.source()).text());
    TextSourceInput.Utf8File fileSource = TextSourceInput.utf8File("raw-formula.txt");
    assertEquals(fileSource, new CellInput.RawFormula(fileSource).source());
  }

  @Test
  void bindsRawFormulaFromTheWireAsADistinctCellInputType() {
    var plan =
        GridGrindJson.readRequest(
            """
            {
              "protocolVersion": "V2",
              "source": { "type": "NEW" },
              "persistence": { "type": "NONE" },
              "steps": [
                {
                  "stepId": "set-raw",
                  "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                  "action": {
                    "type": "SET_CELL",
                    "value": { "type": "RAW_FORMULA", "source": { "type": "INLINE", "text": "LAMBDA(x,x+1)(A1)" } }
                  }
                }
              ]
            }
            """);

    MutationStep step = assertInstanceOf(MutationStep.class, plan.steps().getFirst());
    CellMutationAction.SetCell setCell =
        assertInstanceOf(CellMutationAction.SetCell.class, step.action());
    assertInstanceOf(CellInput.RawFormula.class, setCell.value());
  }

  @Test
  void retainsTheFormulaSpecificProblemCodeThroughJsonBinding() {
    FormulaRequestException normal =
        assertThrows(
            FormulaRequestException.class,
            () -> GridGrindJson.readRequest(requestWithFormula("FORMULA", "=SUM(A1:A2)")));
    FormulaRequestException raw =
        assertThrows(
            FormulaRequestException.class,
            () -> GridGrindJson.readRequest(requestWithFormula("RAW_FORMULA", "A1\u0001+B1")));

    assertEquals(GridGrindProblemCode.INVALID_FORMULA, normal.problemCode());
    assertEquals(GridGrindProblemCode.INVALID_FORMULA_TEXT, raw.problemCode());
    assertEquals(
        "steps[0].action.value.source.text", normal.jsonLocation().jsonPath().orElseThrow());
    assertEquals(
        "steps[0].action.value.source.text", normal.requestProblem().jsonPath().orElseThrow());
    assertTrue(normal.getMessage().contains("must not begin with ="));
  }

  private static String requestWithFormula(String type, String formula) {
    return """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "steps": [
            {
              "stepId": "set-formula",
              "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
              "action": {
                "type": "SET_CELL",
                "value": { "type": "%s", "source": { "type": "INLINE", "text": "%s" } }
              }
            }
          ]
        }
        """
        .formatted(type, formula.replace("\u0001", "\\u0001"));
  }
}
