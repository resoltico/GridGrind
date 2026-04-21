package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.json.GridGrindJson;
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
                    "target": { "type": "CURRENT" },
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
                    "target": { "type": "CURRENT" },
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
                    "target": { "type": "CURRENT" }
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
                    "target": { "type": "CURRENT" },
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
                    "target": { "type": "BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                    "assertion": {
                      "type": "EXPECT_CELL_VALUE",
                      "expectedValue": { "type": "TEXT", "text": "Owner" }
                    }
                    """)));

    assertEquals(1, request.assertionSteps().size());
    assertEquals("assert-owner", request.assertionSteps().getFirst().stepId());
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
                        "target": { "type": "CURRENT" },
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
                        "target": { "type": "BY_ADDRESS", "sheetName": "Budget", "address": "A1" },
                        "query": { "type": "GET_WINDOW" }
                        """)));

    assertEquals(
        "Target selector type 'CURRENT' is not allowed for this step; allowed targets: CellSelector(BY_ADDRESS); TableCellSelector(BY_COLUMN_NAME)",
        wrongMutationTarget.getMessage());
    assertEquals("steps[0].target", wrongMutationTarget.jsonPath());
    assertEquals(
        "Target selector type 'BY_ADDRESS' is not allowed for this step; allowed targets: RangeSelector(RECTANGULAR_WINDOW)",
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
    assertEquals("Unknown type value 'BY_RIDDLE'", unknownTargetType.getMessage());
    assertEquals("steps[0].target", unknownTargetType.jsonPath());
  }

  @Test
  void reportsAmbiguousByNameTargetsWithOneProductOwnedMessage() {
    InvalidRequestShapeException wrongShapeByName =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                        "stepId": "assert-present",
                        "target": { "type": "BY_NAME", "sheetName": "Ops" },
                        "assertion": { "type": "EXPECT_PRESENT" }
                        """)));

    assertEquals(
        "Target selector type 'BY_NAME' has the wrong shape for this step; none of the matching selector families accepted the authored fields. Matching families: NamedRangeSelector(BY_NAME); TableSelector(BY_NAME); PivotTableSelector(BY_NAME); ChartSelector(BY_NAME)",
        wrongShapeByName.getMessage());
    assertEquals("steps[0].target", wrongShapeByName.jsonPath());
  }

  @Test
  void rejectsSharedSelectorWireShapesThatMatchMultipleFamilies() {
    InvalidRequestShapeException ambiguousByName =
        assertThrows(
            InvalidRequestShapeException.class,
            () ->
                GridGrindJson.readRequest(
                    requestWithStepBody(
                        """
                        "stepId": "assert-present",
                        "target": { "type": "BY_NAME", "name": "ExpensePivot" },
                        "assertion": { "type": "EXPECT_PRESENT" }
                        """)));

    assertEquals(
        "Target selector type 'BY_NAME' is ambiguous for this step; the authored object matches multiple selector families: NamedRangeSelector(BY_NAME); TableSelector(BY_NAME); PivotTableSelector(BY_NAME)",
        ambiguousByName.getMessage());
    assertEquals("steps[0].target", ambiguousByName.jsonPath());
  }

  @Test
  void internalSelectorDispatchGuardsStayDeterministic() {
    assertEquals(
        TableSelector.ByName.class,
        WorkbookStepJsonDeserializer.castSelectorType(TableSelector.ByName.class));
    assertEquals(
        "TableSelector(BY_NAME, BY_NAME_ON_SHEET)",
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
          "source": { "type": "NEW" },
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
