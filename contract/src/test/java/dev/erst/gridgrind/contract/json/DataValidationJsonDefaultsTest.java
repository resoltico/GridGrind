package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertFalse;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.node.ObjectNode;

/** Regression coverage for request-side data-validation omission defaults. */
class DataValidationJsonDefaultsTest {
  @Test
  void defaultsDataValidationPayloadsThatOmitValidationBooleans() throws IOException {
    ObjectNode missingAllowBlank =
        requestTree(
            """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "execution": {
            "mode": { "type": "FULL_XSSF" },
            "journal": { "level": "NORMAL" },
            "calculation": {
              "strategy": { "type": "DO_NOT_CALCULATE" },
              "markRecalculateOnOpen": false
            }
          },
          "formulaEnvironment": {
            "externalWorkbooks": [ ],
            "missingWorkbookPolicy": "ERROR",
            "udfToolpacks": [ ]
          },
          "steps": [
            {
              "stepId": "validation-missing-allowBlank",
              "target": {
                "type": "RANGE_BY_RANGE",
                "sheetName": "Intake",
                "range": "A2:A20"
              },
              "action": {
                "type": "SET_DATA_VALIDATION",
                "validation": {
                  "rule": {
                    "type": "EXPLICIT_LIST",
                    "values": [ "Queued" ]
                  },
                  "suppressDropDownArrow": false
                }
              }
            }
          ]
        }
        """);
    ObjectNode missingSuppressDropDownArrow =
        requestTree(
            """
        {
          "protocolVersion": "V2",
          "source": { "type": "NEW" },
          "persistence": { "type": "NONE" },
          "execution": {
            "mode": { "type": "FULL_XSSF" },
            "journal": { "level": "NORMAL" },
            "calculation": {
              "strategy": { "type": "DO_NOT_CALCULATE" },
              "markRecalculateOnOpen": false
            }
          },
          "formulaEnvironment": {
            "externalWorkbooks": [ ],
            "missingWorkbookPolicy": "ERROR",
            "udfToolpacks": [ ]
          },
          "steps": [
            {
              "stepId": "validation-missing-suppressDropDownArrow",
              "target": {
                "type": "RANGE_BY_RANGE",
                "sheetName": "Intake",
                "range": "A2:A20"
              },
              "action": {
                "type": "SET_DATA_VALIDATION",
                "validation": {
                  "rule": {
                    "type": "EXPLICIT_LIST",
                    "values": [ "Queued" ]
                  },
                  "allowBlank": false
                }
              }
            }
          ]
        }
        """);

    assertDefaultBooleans(missingAllowBlank);
    assertDefaultBooleans(missingSuppressDropDownArrow);
  }

  private static ObjectNode requestTree(String requestJson) throws IOException {
    return GridGrindJsonOutput.requestTree(
        GridGrindJson.readRequest(requestJson.getBytes(StandardCharsets.UTF_8)));
  }

  private static void assertDefaultBooleans(ObjectNode request) {
    assertFalse(
        request
            .path("steps")
            .path(0)
            .path("action")
            .path("validation")
            .path("allowBlank")
            .booleanValue());
    assertFalse(
        request
            .path("steps")
            .path(0)
            .path("action")
            .path("validation")
            .path("suppressDropDownArrow")
            .booleanValue());
  }
}
