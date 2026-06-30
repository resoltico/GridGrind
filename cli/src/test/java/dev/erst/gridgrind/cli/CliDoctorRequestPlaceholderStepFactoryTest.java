package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import java.io.IOException;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.json.JsonMapper;
import tools.jackson.databind.node.IntNode;
import tools.jackson.databind.node.ObjectNode;

/** Focused branch coverage for synthetic placeholder-step generation in doctor preflight. */
class CliDoctorRequestPlaceholderStepFactoryTest {
  private static final JsonMapper JSON = JsonMapper.builder().build();

  @Test
  void authoredExecutionModeTypeDefaultsAndPreservesScalarAuthoredValues() {
    ObjectNode root = JSON.createObjectNode();

    assertEquals(
        "FULL_XSSF", CliDoctorRequestPlaceholderStepFactory.authoredExecutionModeType(root));

    root.withObject("execution").withObject("mode").putNull("type");
    assertEquals(
        "FULL_XSSF", CliDoctorRequestPlaceholderStepFactory.authoredExecutionModeType(root));

    root.withObject("execution").withObject("mode").put("type", 123);
    assertEquals("123", CliDoctorRequestPlaceholderStepFactory.authoredExecutionModeType(root));

    root.withObject("execution").withObject("mode").put("type", true);
    assertEquals("true", CliDoctorRequestPlaceholderStepFactory.authoredExecutionModeType(root));

    root.withObject("execution").withObject("mode").set("type", JSON.createArrayNode().add("x"));
    assertEquals(
        "FULL_XSSF", CliDoctorRequestPlaceholderStepFactory.authoredExecutionModeType(root));

    root.withObject("execution").withObject("mode").put("type", "EVENT_READ");
    assertEquals(
        "EVENT_READ", CliDoctorRequestPlaceholderStepFactory.authoredExecutionModeType(root));
  }

  @Test
  void placeholderStepNodeDefaultsToQueryForNonObjectSteps() {
    ObjectNode placeholder =
        CliDoctorRequestPlaceholderStepFactory.placeholderStepNode(
            IntNode.valueOf(7), 2, "FULL_XSSF");

    assertEquals("gridgrind-doctor-step-2", placeholder.path("stepId").asString());
    assertEquals("GET_WORKBOOK_SUMMARY", placeholder.path("query").path("type").asString());
  }

  @Test
  void placeholderStepNodePreservesValidMutationStepIds() throws IOException {
    ObjectNode placeholder =
        CliDoctorRequestPlaceholderStepFactory.placeholderStepNode(
            objectNode(
                """
                {
                  "stepId": "keep-me",
                  "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
                  "action": { "type": "SET_SHEET_ZOOM", "zoomPercent": 125 }
                }
                """),
            0,
            "EVENT_READ");

    assertEquals("keep-me", placeholder.path("stepId").asString());
    assertEquals("ENSURE_SHEET", placeholder.path("action").path("type").asString());
    assertEquals("GridGrindDoctorStep", placeholder.path("target").path("name").asString());
  }

  @Test
  void placeholderStepNodeHandlesAssertionQueryAndStreamingWriteFallbacks() throws IOException {
    ObjectNode assertionPlaceholder =
        CliDoctorRequestPlaceholderStepFactory.placeholderStepNode(
            objectNode(
                """
                {
                  "stepId": "",
                  "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
                  "assertion": { "type": "EXPECT_SHEET_PRESENT" }
                }
                """),
            1,
            "STREAMING_WRITE");
    ObjectNode numericStepIdMutationPlaceholder =
        CliDoctorRequestPlaceholderStepFactory.placeholderStepNode(
            objectNode(
                """
                {
                  "stepId": 7,
                  "target": { "type": "SHEET_BY_NAME", "name": "Budget" },
                  "action": { "type": "SET_SHEET_ZOOM", "zoomPercent": 125 }
                }
                """),
            2,
            "STREAMING_WRITE");
    ObjectNode queryPlaceholder =
        CliDoctorRequestPlaceholderStepFactory.placeholderStepNode(
            objectNode(
                """
                {
                  "stepId": "summary",
                  "target": { "type": "WORKBOOK_CURRENT" },
                  "query": { "type": "GET_WORKBOOK_SUMMARY" }
                }
                """),
            3,
            "EVENT_READ");
    ObjectNode defaultMutationPlaceholder =
        CliDoctorRequestPlaceholderStepFactory.placeholderStepNode(
            objectNode("{ }"), 4, "STREAMING_WRITE");

    assertEquals("gridgrind-doctor-step-1", assertionPlaceholder.path("stepId").asString());
    assertTrue(assertionPlaceholder.has("assertion"));
    assertEquals(
        "gridgrind-doctor-step-2", numericStepIdMutationPlaceholder.path("stepId").asString());
    assertEquals(
        "ENSURE_SHEET", numericStepIdMutationPlaceholder.path("action").path("type").asString());
    assertEquals("summary", queryPlaceholder.path("stepId").asString());
    assertTrue(queryPlaceholder.has("query"));
    assertEquals("gridgrind-doctor-step-4", defaultMutationPlaceholder.path("stepId").asString());
    assertEquals("ENSURE_SHEET", defaultMutationPlaceholder.path("action").path("type").asString());
  }

  private static ObjectNode objectNode(String json) throws IOException {
    JsonNode node = JSON.readTree(json);
    return (ObjectNode) node;
  }
}
