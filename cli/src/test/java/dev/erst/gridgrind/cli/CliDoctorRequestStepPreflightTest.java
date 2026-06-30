package dev.erst.gridgrind.cli;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import java.io.IOException;
import org.junit.jupiter.api.Test;
import tools.jackson.databind.JsonNode;
import tools.jackson.databind.json.JsonMapper;
import tools.jackson.databind.node.ObjectNode;

/** Focused coverage for step-level doctor preflight batching and path rebasing. */
class CliDoctorRequestStepPreflightTest {
  private static final JsonMapper JSON = JsonMapper.builder().build();

  @Test
  void fromReturnsCleanResultWhenStepsArrayIsMissing() {
    CliDoctorRequestStepPreflight.StepPreflight preflight =
        CliDoctorRequestStepPreflight.from(
            JSON.createObjectNode(), ProblemContextRequestSurfaces.RequestInput.standardInput());

    assertEquals(java.util.List.of(), preflight.problems());
    assertFalse(preflight.usesSyntheticValues());
    assertTrue(preflight.summaryTrustworthy());
  }

  @Test
  void fromRebasesMalformedStepShapeProblemsToTheAuthoredIndex() throws IOException {
    CliDoctorRequestStepPreflight.StepPreflight preflight =
        CliDoctorRequestStepPreflight.from(
            objectNode(
                """
                {
                  "steps": [
                    {
                      "stepId": "summary",
                      "target": { "type": "WORKBOOK_CURRENT" },
                      "query": { "type": "GET_WORKBOOK_SUMMARY" }
                    },
                    {
                      "stepId": "broken-query",
                      "target": { "type": "WORKBOOK_CURRENT" },
                      "query": { }
                    }
                  ]
                }
                """),
            ProblemContextRequestSurfaces.RequestInput.requestFile("/tmp/request.json"));

    ProblemContext.ReadRequest context =
        assertInstanceOf(
            ProblemContext.ReadRequest.class, preflight.problems().getFirst().context());

    assertEquals(1, preflight.problems().size());
    assertEquals(java.util.Optional.of("steps[1].query.type"), context.jsonPath());
    assertEquals("Missing required field 'type'", preflight.problems().getFirst().message());
    assertTrue(preflight.usesSyntheticValues());
    assertFalse(preflight.summaryTrustworthy());
  }

  @Test
  void rebaseMessagePathLeavesStepZeroMessagesUntouched() {
    assertEquals(
        "Missing required field 'steps[0].query.type'",
        CliDoctorRequestStepPreflight.rebaseMessagePath(
            "Missing required field 'steps[0].query.type'", 0));
  }

  @Test
  void rebaseMessagePathRewritesEmbeddedStepZeroPathsForLaterIndexes() {
    assertEquals(
        "Missing required field 'steps[2].query.type'",
        CliDoctorRequestStepPreflight.rebaseMessagePath(
            "Missing required field 'steps[0].query.type'", 2));
  }

  @Test
  void rebaseMessagePathLeavesMessagesWithoutEmbeddedJsonPathsUntouched() {
    assertEquals(
        "plain problem", CliDoctorRequestStepPreflight.rebaseMessagePath("plain problem", 4));
  }

  private static ObjectNode objectNode(String json) throws IOException {
    JsonNode node = JSON.readTree(json);
    return (ObjectNode) node;
  }
}
