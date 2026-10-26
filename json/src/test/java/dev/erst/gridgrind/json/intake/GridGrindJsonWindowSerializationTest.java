package dev.erst.gridgrind.json.intake;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.json.codec.GridGrindJson;
import dev.erst.gridgrind.json.codec.GridGrindJsonMapperSupport;
import dev.erst.gridgrind.json.codec.GridGrindJsonOutput;
import java.util.Optional;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;

/**
 * Locks the sparse GET_WINDOW wire default so false is omitted unless dense output is requested.
 */
final class GridGrindJsonWindowSerializationTest {
  @Test
  @SuppressWarnings(
      "NullOptional") // Deliberately proves the Jackson creator normalizes absent input.
  void getWindowCreatorNormalizesEachOmittedAxisAtTheDeserializationBoundary() {
    SheetIntrospectionQuery.GetWindow defaults =
        GridGrindJsonMapperSupport.JSON_MAPPER.readValue(
            "{\"type\":\"GET_WINDOW\"}", SheetIntrospectionQuery.GetWindow.class);
    SheetIntrospectionQuery.GetWindow dense =
        new SheetIntrospectionQuery.GetWindow(Optional.empty(), true);

    assertEquals(Optional.empty(), defaults.projection());
    assertFalse(defaults.includeBlanks());
    assertTrue(dense.includeBlanks());
  }

  @Test
  void sparseWindowRequestsOmitIncludeBlanksUntilDenseReadbackIsRequested() {
    WorkbookPlan sparseWindowRequest =
        assertDoesNotThrow(
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": 1,
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "steps": [
                        {
                          "stepId": "window",
                          "target": {
                            "type": "RANGE_RECTANGULAR_WINDOW",
                            "sheetName": "Budget",
                            "topLeftAddress": "A1",
                            "rowCount": 4,
                            "columnCount": 3
                          },
                          "query": { "type": "GET_WINDOW" }
                        }
                      ]
                    }
                    """));
    WorkbookPlan denseWindowRequest =
        assertDoesNotThrow(
            () ->
                GridGrindJson.readRequest(
                    """
                    {
                      "protocolVersion": 1,
                      "source": { "type": "NEW" },
                      "persistence": { "type": "NONE" },
                      "steps": [
                        {
                          "stepId": "window",
                          "target": {
                            "type": "RANGE_RECTANGULAR_WINDOW",
                            "sheetName": "Budget",
                            "topLeftAddress": "A1",
                            "rowCount": 4,
                            "columnCount": 3
                          },
                          "query": { "type": "GET_WINDOW", "includeBlanks": true }
                        }
                      ]
                    }
                    """));

    JsonNode sparseWindowQuery =
        GridGrindJsonOutput.requestTree(sparseWindowRequest).path("steps").get(0).path("query");
    JsonNode denseWindowQuery =
        GridGrindJsonOutput.requestTree(denseWindowRequest).path("steps").get(0).path("query");

    assertEquals("GET_WINDOW", sparseWindowQuery.path("type").stringValue());
    assertFalse(sparseWindowQuery.has("includeBlanks"));
    assertTrue(denseWindowQuery.path("includeBlanks").booleanValue());
  }
}
