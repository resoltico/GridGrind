package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.nio.charset.StandardCharsets;
import org.junit.jupiter.api.Test;

/** Covers the explicit request JSON contract with no request-only default expansion. */
class GridGrindJsonRequestContractTest {
  @Test
  void requestRequiresExplicitProtocolVersion() {
    InvalidRequestException exception =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
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
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertTrue(exception.getMessage().contains("protocolVersion"));
  }

  @Test
  void requestRequiresExplicitPersistenceAndSteps() {
    InvalidRequestException exception =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V1",
                      "source": { "type": "NEW" },
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
                      }
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertTrue(
        exception.getMessage().contains("persistence") || exception.getMessage().contains("steps"));
  }

  @Test
  void requestRequiresExplicitTopLevelExecutionAndFormulaEnvironment() {
    InvalidRequestException exception =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": "V1",
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "steps": []
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertTrue(exception.getMessage().contains("execution"));
  }

  @Test
  void requestRejectsSparsePrintLayoutPayloadsThatPreviouslyReliedOnImplicitDefaults() {
    InvalidRequestException exception =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
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
                      "steps": [ {
                        "stepId": "layout",
                        "target": { "type": "SHEET_BY_NAME", "name": "Sheet1" },
                        "action": {
                          "type": "SET_PRINT_LAYOUT",
                          "printLayout": {
                            "printArea": { "type": "NONE" },
                            "orientation": "PORTRAIT",
                            "scaling": { "type": "AUTOMATIC" },
                            "repeatingRows": { "type": "NONE" },
                            "repeatingColumns": { "type": "NONE" },
                            "header": {
                              "left": { "type": "INLINE", "text": "" },
                              "center": { "type": "INLINE", "text": "" },
                              "right": { "type": "INLINE", "text": "" }
                            },
                            "footer": {
                              "left": { "type": "INLINE", "text": "" },
                              "center": { "type": "INLINE", "text": "" },
                              "right": { "type": "INLINE", "text": "" }
                            }
                          }
                        }
                      } ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertTrue(exception.getMessage().contains("setup"));
  }

  @Test
  void requestRejectsSparseChartPayloadsThatPreviouslyReliedOnImplicitDefaults() {
    InvalidRequestException exception =
        assertThrows(
            InvalidRequestException.class,
            () ->
                GridGrindJson.readRequest(
                    """
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
                      "steps": [ {
                        "stepId": "chart",
                        "target": { "type": "DRAWING_OBJECT_BY_NAME_ON_SHEET", "sheetName": "Sheet1", "name": "RevenueChart" },
                        "action": {
                          "type": "SET_CHART",
                          "chart": {
                            "name": "RevenueChart",
                            "anchor": {
                              "type": "TWO_CELL",
                              "start": { "columnIndex": 0, "rowIndex": 0, "dx": 0, "dy": 0 },
                              "end": { "columnIndex": 8, "rowIndex": 14, "dx": 0, "dy": 0 },
                              "behavior": "MOVE_AND_RESIZE"
                            },
                            "plots": [ {
                              "type": "LINE",
                              "varyColors": false,
                              "grouping": "STANDARD",
                              "axes": [
                                {
                                  "kind": "CATEGORY",
                                  "position": "BOTTOM",
                                  "crosses": "AUTO_ZERO",
                                  "visible": true
                                },
                                {
                                  "kind": "VALUE",
                                  "position": "LEFT",
                                  "crosses": "AUTO_ZERO",
                                  "visible": true
                                }
                              ],
                              "series": [ {
                                "title": { "type": "NONE" },
                                "categories": { "type": "REFERENCE", "reference": "Sheet1!$A$2:$A$5" },
                                "values": { "type": "REFERENCE", "reference": "Sheet1!$B$2:$B$5" }
                              } ]
                            } ]
                          }
                        }
                      } ]
                    }
                    """
                        .getBytes(StandardCharsets.UTF_8)));

    assertTrue(exception.getMessage().contains("Missing required field 'from'"));
  }
}
