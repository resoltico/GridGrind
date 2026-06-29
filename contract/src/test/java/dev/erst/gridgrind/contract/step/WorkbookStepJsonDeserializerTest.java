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
import java.util.Optional;
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
    InvalidRequestShapeException unknownField =
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
                    """)));
    InvalidRequestShapeException missingStepId =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                    "target": { "type": "WORKBOOK_CURRENT" },
                    "query": { "type": "GET_WORKBOOK_SUMMARY" }
                    """)));
    InvalidRequestShapeException missingTarget =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                    "stepId": "bad",
                    "query": { "type": "GET_WORKBOOK_SUMMARY" }
                    """)));
    InvalidRequestShapeException missingStepPayload =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                    "stepId": "bad",
                    "target": { "type": "WORKBOOK_CURRENT" }
                    """)));
    InvalidRequestShapeException nonStringStepId =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                    "stepId": 3,
                    "target": { "type": "WORKBOOK_CURRENT" },
                    "query": { "type": "GET_WORKBOOK_SUMMARY" }
                    """)));

    assertEquals("Unknown field 'legacy'", unknownField.getMessage());
    assertEquals(Optional.of("steps[0].legacy"), unknownField.jsonPath());
    assertEquals("Missing required field 'stepId'", missingStepId.getMessage());
    assertEquals(Optional.of("steps[0].stepId"), missingStepId.jsonPath());
    assertEquals("Missing required field 'target'", missingTarget.getMessage());
    assertEquals(Optional.of("steps[0].target"), missingTarget.jsonPath());
    assertEquals(
        "Each step must contain exactly one of 'action', 'assertion', or 'query'",
        missingStepPayload.getMessage());
    assertEquals(Optional.of("steps[0]"), missingStepPayload.jsonPath());
    assertEquals("Field 'stepId' must be a string", nonStringStepId.getMessage());
    assertEquals(Optional.of("steps[0].stepId"), nonStringStepId.jsonPath());
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
    assertEquals(Optional.of("steps[0].target.type"), wrongMutationTarget.jsonPath());
    assertEquals(
        "Target selector type 'CELL_BY_ADDRESS' is not allowed for this step; allowed targets: RangeSelector(RANGE_RECTANGULAR_WINDOW)",
        wrongInspectionTarget.getMessage());
    assertEquals(Optional.of("steps[0].target.type"), wrongInspectionTarget.jsonPath());
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
    assertEquals(Optional.of("steps[0].target"), nonObjectTarget.jsonPath());
    assertEquals("Missing required field 'type'", missingTargetType.getMessage());
    assertEquals(Optional.of("steps[0].target.type"), missingTargetType.jsonPath());
    assertEquals("Field 'type' must be a string", nonStringTargetType.getMessage());
    assertEquals(Optional.of("steps[0].target.type"), nonStringTargetType.jsonPath());
    assertEquals(
        "Unknown target selector type 'BY_RIDDLE'; allowed targets: WorkbookSelector(WORKBOOK_CURRENT)",
        unknownTargetType.getMessage());
    assertEquals(Optional.of("steps[0].target.type"), unknownTargetType.jsonPath());
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
    assertEquals(Optional.of("steps[0].action.zoomPercent"), invalidAction.jsonPath());
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
    assertEquals(Optional.of("steps[0].target.type"), legacyTargetType.jsonPath());
  }

  @Test
  void reportsWrongShapeWithinOneSelectorFamilyAgainstTheTargetField() {
    InvalidRequestShapeException wrongShapeByName =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                        "stepId": "assert-table-present",
                        "target": { "type": "TABLE_BY_NAME", "sheetName": "Ops" },
                        "assertion": { "type": "EXPECT_TABLE_PRESENT" }
                        """)));

    assertEquals("Missing required field 'name'", wrongShapeByName.getMessage());
    assertEquals(Optional.of("steps[0].target.name"), wrongShapeByName.jsonPath());
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
        WorkbookStepJsonTargetSupport.castSelectorType(TableSelector.ByName.class));
    assertEquals(
        "TableSelector(TABLE_BY_NAME, TABLE_BY_NAME_ON_SHEET)",
        WorkbookStepJsonTargetSupport.selectorFamilySummary(
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
            "mode": {"type": "FULL_XSSF"},
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
