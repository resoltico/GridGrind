package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.json.GridGrindJson;
import dev.erst.gridgrind.contract.json.InvalidRequestException;
import dev.erst.gridgrind.contract.json.InvalidRequestShapeException;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SelectorJsonSupport;
import dev.erst.gridgrind.contract.selector.TableSelector;
import java.nio.charset.StandardCharsets;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct parser coverage for the canonical step envelope deserializer. */
class WorkbookStepJsonDeserializerTest {
  @Test
  void rejectsUnknownAndMissingStepFieldsWithProductOwnedMessages() {
    assertEquals(
        "steps entries must be JSON objects",
        assertThrows(
                InvalidRequestShapeException.class,
                () ->
                    GridGrindJson.readRequest(
                        """
                        {
                          "source": { "type": "NEW" },
                          "steps": [3]
                        }
                        """
                            .getBytes(StandardCharsets.UTF_8)))
            .getMessage());
    assertEquals(
        "Unknown field 'legacy'",
        assertThrows(
                InvalidRequestShapeException.class,
                () ->
                    GridGrindJson.readRequest(
                        requestWithStepBody(
                            """
                    "stepId": "bad",
                    "target": { "type": "WORKBOOK_CURRENT" },
                    "query": { "type": "GET_WORKBOOK_SUMMARY" },
                    "legacy": true
                    """)))
            .getMessage());
    assertEquals(
        "Missing required field 'stepId'",
        assertThrows(
                InvalidRequestShapeException.class,
                () ->
                    GridGrindJson.readRequest(
                        requestWithStepBody(
                            """
                    "target": { "type": "WORKBOOK_CURRENT" },
                    "query": { "type": "GET_WORKBOOK_SUMMARY" }
                    """)))
            .getMessage());
    assertEquals(
        "Missing required field 'target'",
        assertThrows(
                InvalidRequestShapeException.class,
                () ->
                    GridGrindJson.readRequest(
                        requestWithStepBody(
                            """
                    "stepId": "bad",
                    "query": { "type": "GET_WORKBOOK_SUMMARY" }
                    """)))
            .getMessage());
    assertEquals(
        "Each step must contain exactly one of 'action', 'assertion', or 'query'",
        assertThrows(
                InvalidRequestShapeException.class,
                () ->
                    GridGrindJson.readRequest(
                        requestWithStepBody(
                            """
                    "stepId": "bad",
                    "target": { "type": "WORKBOOK_CURRENT" }
                    """)))
            .getMessage());
    assertEquals(
        "Field 'stepId' must be a string",
        assertThrows(
                InvalidRequestShapeException.class,
                () ->
                    GridGrindJson.readRequest(
                        requestWithStepBody(
                            """
                    "stepId": 3,
                    "target": { "type": "WORKBOOK_CURRENT" },
                    "query": { "type": "GET_WORKBOOK_SUMMARY" }
                    """)))
            .getMessage());
  }

  @Test
  void readsAssertionStepsThroughTheCanonicalEnvelope() {
    var request =
        assertDoesNotThrow(
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                    "stepId": "assert-owner",
                    "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                    "assertion": {
                      "type": "EXPECT_CELL_VALUE",
                      "expectedValue": { "type": "TEXT", "text": "Owner" }
                    }
                    """)));

    assertEquals(1, request.stepPartition().assertions().size());
    assertEquals("assert-owner", request.stepPartition().assertions().getFirst().stepId());
  }

  @Test
  void reportsDisallowedSelectorTypesAgainstTheTargetField() {
    InvalidRequestShapeException wrongMutationTarget =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                        "stepId": "set-cell",
                        "target": { "type": "WORKBOOK_CURRENT" },
                        "action": {
                          "type": "SET_CELL",
                          "value": {
                            "type": "TEXT",
                            "source": { "type": "INLINE", "text": "Owner" }
                          }
                        }
                        """)));
    InvalidRequestShapeException wrongInspectionTarget =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                        "stepId": "window",
                        "target": { "type": "CELL_BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                        "query": { "type": "GET_WINDOW" }
                        """)));

    assertEquals(
        "Target selector type 'WORKBOOK_CURRENT' is not allowed for this step; allowed targets: CellSelector(CELL_BY_ADDRESS); TableCellSelector(TABLE_CELL_BY_COLUMN_NAME)",
        wrongMutationTarget.getMessage());
    assertEquals("steps[0].target", wrongMutationTarget.jsonPath());
    assertEquals(
        "Target selector type 'CELL_BY_ADDRESS' is not allowed for this step; allowed targets: RangeSelector(RANGE_RECTANGULAR_WINDOW)",
        wrongInspectionTarget.getMessage());
    assertEquals("steps[0].target", wrongInspectionTarget.jsonPath());
  }

  @Test
  void reportsMalformedTargetShapesAgainstTheTargetField() {
    InvalidRequestShapeException nonObjectTarget =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                        "stepId": "bad-target",
                        "target": 3,
                        "query": { "type": "GET_WORKBOOK_SUMMARY" }
                        """)));
    InvalidRequestShapeException missingTargetType =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                        "stepId": "bad-target",
                        "target": { "sheetName": "Ops" },
                        "query": { "type": "GET_SHEET_SUMMARY" }
                        """)));
    InvalidRequestShapeException nonStringTargetType =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                        "stepId": "bad-target",
                        "target": { "type": 7, "sheetName": "Ops" },
                        "query": { "type": "GET_SHEET_SUMMARY" }
                        """)));
    InvalidRequestShapeException unknownTargetType =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                        "stepId": "bad-target",
                        "target": { "type": "BY_RIDDLE" },
                        "query": { "type": "GET_WORKBOOK_SUMMARY" }
                        """)));

    assertEquals("Field 'target' must be a JSON object", nonObjectTarget.getMessage());
    assertEquals("steps[0].target", nonObjectTarget.jsonPath());
    assertEquals("Missing required field 'type'", missingTargetType.getMessage());
    assertEquals("steps[0].target", missingTargetType.jsonPath());
    assertEquals("Field 'type' must be a string", nonStringTargetType.getMessage());
    assertEquals("steps[0].target", nonStringTargetType.jsonPath());
    assertEquals(
        "Unknown target selector type 'BY_RIDDLE'; allowed targets: WorkbookSelector(WORKBOOK_CURRENT)",
        unknownTargetType.getMessage());
    assertEquals("steps[0].target", unknownTargetType.jsonPath());
  }

  @Test
  void wrapsIllegalArgumentActionPayloadsAgainstTheActionField() {
    InvalidRequestException invalidAction =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                        "stepId": "zoom-too-far",
                        "target": { "type": "SHEET_BY_NAME", "sheetName": "Budget" },
                        "action": { "type": "SET_SHEET_ZOOM", "zoomPercent": 401 }
                        """)));

    assertEquals(
        "zoomPercent must be between 10 and 400 inclusive: 401", invalidAction.getMessage());
    assertEquals("steps[0].action", invalidAction.jsonPath());
  }

  @Test
  void reportsLegacyGenericTargetIdsWithFamilySpecificGuidance() {
    InvalidRequestShapeException legacyTargetType =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                        "stepId": "bad-target",
                        "target": { "type": "BY_NAME" },
                        "query": { "type": "GET_WORKBOOK_SUMMARY" }
                        """)));

    assertEquals(
        "Unknown target selector type 'BY_NAME'; target selector ids are family-specific; allowed targets: WorkbookSelector(WORKBOOK_CURRENT)",
        legacyTargetType.getMessage());
    assertEquals("steps[0].target", legacyTargetType.jsonPath());
  }

  @Test
  void reportsWrongShapeWithinOneSelectorFamilyAgainstTheTargetField() {
    InvalidRequestException wrongShapeByName =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                        "stepId": "assert-table-present",
                        "target": { "type": "TABLE_BY_NAME", "sheetName": "Ops" },
                        "assertion": { "type": "EXPECT_TABLE_PRESENT" }
                        """)));

    assertEquals("Missing required field 'name'", wrongShapeByName.getMessage());
    assertEquals("steps[0].target", wrongShapeByName.jsonPath());
  }

  @Test
  void readsFamilySpecificPresenceAssertionsWithoutSharedSelectorAmbiguity() {
    assertDoesNotThrow(
        () ->
            GridGrindJson.readRequest(
                requestWithStepBody(
                    """
                    "stepId": "assert-pivot-present",
                    "target": { "type": "PIVOT_TABLE_BY_NAME", "name": "ExpensePivot" },
                    "assertion": { "type": "EXPECT_PIVOT_TABLE_PRESENT" }
                    """)));
  }

  @Test
  void internalSelectorDispatchGuardsStayDeterministic() {
    assertEquals(
        TableSelector.ByName.class,
        WorkbookStepJsonDeserializer.castSelectorType(TableSelector.ByName.class));
    assertEquals(
        "TableSelector(TABLE_BY_NAME, TABLE_BY_NAME_ON_SHEET)",
        WorkbookStepJsonDeserializer.selectorFamilySummary(
            List.of(
                (Class<? extends Selector>) TableSelector.ByName.class,
                TableSelector.ByNameOnSheet.class)));
    assertEquals("Selector", SelectorJsonSupport.familyName(Selector.class));
    assertEquals("TableSelector", SelectorJsonSupport.familyName(TableSelector.ByName.class));
  }

  private static byte[] requestWithStepBody(String stepBody) {
    return ("""
        {
          "protocolVersion": "V1",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "execution": {
            "mode": { "readMode": "FULL_XSSF", "writeMode": "FULL_XSSF" },
            "journal": { "level": "NORMAL" },
            "calculation": {
              "strategy": { "type": "DO_NOT_CALCULATE" },
              "markRecalculateOnOpen": false
            }
          },
          "formulaEnvironment": {
            "externalWorkbooks": [],
            "missingWorkbookPolicy": "ERROR",
            "udfToolpacks": []
          },
          "steps": [
            {
        """
            + stepBody
            + """
            }
          ]
        }
        """)
        .getBytes(StandardCharsets.UTF_8);
  }
}
