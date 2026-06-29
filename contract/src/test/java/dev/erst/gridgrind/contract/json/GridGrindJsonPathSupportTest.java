package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.json.JsonMapper;

/** Direct coverage for dotted-and-indexed JSON-path traversal helpers. */
class GridGrindJsonPathSupportTest {
  private static final JsonMapper JSON_MAPPER = JsonMapper.builder().build();

  @Test
  void findsRootNestedAndExplicitNullPaths() throws IOException {
    JsonNode objectNode =
        readTree(
            """
            {
              "protocolVersion": "V1",
              "steps": [
                {
                  "target": {
                    "name": null,
                    "matrix": [[1], [2]]
                  }
                }
              ]
            }
            """);
    JsonNode arrayNode =
        readTree(
            """
            [
              { "value": 1 }
            ]
            """);

    assertTrue(GridGrindJsonPathSupport.pathExists(objectNode, "protocolVersion"));
    assertTrue(GridGrindJsonPathSupport.pathExists(objectNode, "steps[0].target.name"));
    assertTrue(GridGrindJsonPathSupport.pathExists(objectNode, "steps[0].target.matrix[1][0]"));
    assertTrue(GridGrindJsonPathSupport.pathExists(arrayNode, "[0].value"));
  }

  @Test
  void returnsFalseWhenPropertiesOrIndexesAreMissing() throws IOException {
    JsonNode node =
        readTree(
            """
            {
              "steps": [
                {
                  "target": {
                    "matrix": [[1]]
                  }
                }
              ]
            }
            """);

    assertFalse(GridGrindJsonPathSupport.pathExists(node, "steps[0].target.name"));
    assertFalse(GridGrindJsonPathSupport.pathExists(node, "steps[1].target"));
    assertFalse(GridGrindJsonPathSupport.pathExists(node, "steps[0].target.matrix[0][1]"));
  }

  @Test
  void returnsFalseForMalformedBracketPaths() throws IOException {
    JsonNode node =
        readTree(
            """
            {
              "steps": [ { "target": { "name": "Ops" } } ]
            }
            """);

    assertFalse(GridGrindJsonPathSupport.pathExists(node, "steps[0"));
  }

  private static JsonNode readTree(String json) throws IOException {
    return JSON_MAPPER.readTree(json);
  }
}
